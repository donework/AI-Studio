#!/usr/bin/env python3
"""
A stand-in for a SharePoint Server, for trying the Search SharePoint tool in the app without an
intranet. It answers the search API with the made-up fixtures next to this file and serves the
pages and PDFs those hits point to. It needs Python 3 and nothing else.

It is no SharePoint: it asks for no sign-in, ignores what is searched for, and knows one answer.
Whether a real SharePoint accepts the sign-in of AI Studio only a real SharePoint can tell.

Setup, once:

  1. AI Studio refuses loopback addresses for tools, so the server listens on the address of this
     computer in the local network. Start it once without a certificate to learn that address:

         python3 sharepoint-mock-server.py --print-address

  2. Give that address a name. Add the line the command printed to /etc/hosts, such as:

         192.168.1.23 intranet.test

  3. AI Studio accepts HTTPS only, with a certificate this computer trusts. Create one with mkcert
     (https://github.com/FiloSottile/mkcert), outside of the repository:

         mkcert -install
         mkcert -cert-file ~/intranet.test.pem -key-file ~/intranet.test-key.pem intranet.test

Run:

    python3 sharepoint-mock-server.py --cert ~/intranet.test.pem --key ~/intranet.test-key.pem

Then, in AI Studio:

  - Search SharePoint, SharePoint Base URL:   https://intranet.test:8443/
  - Read Web Page, Allowed Private Hosts:     intranet.test
  - a provider with High confidence

What it answers:

  - Every search returns the hits of sharepoint-search.xml, with their hosts replaced by this
    server, so the hits can be opened.
  - The second page of a search returns the two hits of sharepoint-search-page-2.xml, which
    appear nowhere on the first page.
  - A search whose words contain "nothing", and every page after the second, returns the answer
    without hits, sharepoint-search-empty.xml.
  - With --sign-in-page, the search API answers with a web page instead, the way a SharePoint
    does which wants a sign-in AI Studio cannot give.
"""

import argparse
import re
import socket
import ssl
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlsplit

FIXTURES = Path(__file__).resolve().parent

# The hosts the hits of the fixture lie on. A SharePoint farm serves several, the stand-in one:
FIXTURE_HOSTS = re.compile(r"https://(?:intranet|docs|teams)\.example\.org")

# The answers by the row a page starts with. The tool asks for ten hits per page:
FIXTURE_BY_START_ROW = {
    "0": "sharepoint-search.xml",
    "10": "sharepoint-search-page-2.xml",
}

PDFS = {
    "/Documents/Travel Expense Guideline.pdf": "sharepoint-travel-expense-guideline.pdf",
    "/sites/Standards/Shared Documents/Visitor Registration.pdf": "sharepoint-visitor-registration.pdf",
}

PAGES = {
    "/": ("Finance Team Site", [
        "This is the team site of the finance department of Example Corp.",
        "Here you find the template for travel requests, the overview of the expenses of 2025, and the travel budget of 2026.",
        "A travel request is handed in at least ten working days before the trip starts. The head of your department approves it, and the finance department books it afterward.",
        "The overview of the expenses is updated on the first working day of every month. Questions about a single entry go to the finance department, room A 1.04, extension 2300.",
    ]),
    "/News/Pages/travel-expense-rules.aspx": ("New travel expense rules from January", [
        "From January 1, new rules apply to the travel expenses of all staff of Example Corp.",
        "The daily allowance rises to 28 EUR for a full day and to 14 EUR for the day of arrival and the day of departure. Receipts are handed in within four weeks after the trip, no longer within eight.",
        "A rail journey is preferred over a flight for every travel below 600 km. Exceptions need the approval of the head of the department before the trip is booked.",
        "The complete rules are in the Travel Expense Guideline in the document library. Questions go to the travel desk, room B 2.17, extension 4711.",
    ]),
    "/Health/Pages/travel-vaccinations.aspx": ("Vaccinations before a business trip", [
        "The company doctor advises on vaccinations before every travel outside of Europe.",
        "Make an appointment at least six weeks before the trip starts, because some vaccinations need several doses. The consultation takes about twenty minutes and happens during working hours.",
        "Example Corp pays for every vaccination the company doctor recommends for the trip. Bring your vaccination record to the appointment.",
        "The company doctor is in building C, room C 0.12, on Tuesdays and Thursdays between 9:00 and 15:00, extension 1120.",
    ]),
    "/Facilities/Pages/pool-cars.aspx": ("Pool cars and car sharing", [
        "Example Corp keeps six pool cars at the main site, four of them electric.",
        "A pool car is booked for a travel of up to three days through the facility desk. For a longer trip, rent a car through the travel desk instead.",
        "Hand the key back at the gate after the trip and note the mileage in the logbook in the glove box. Charge an electric car at one of the stations in front of building A before you return it.",
        "The facility desk is in building A, room A 0.03, extension 3400.",
    ]),
    "/Pages/canteen.aspx": ("Canteen menu", [
        "The canteen is open from Monday to Friday between 11:30 and 14:00.",
        "This week: lentil soup on Monday, vegetable lasagne on Tuesday, fish with potatoes on Wednesday, chili sin carne on Thursday, and pancakes on Friday.",
        "The canteen stays closed during the travel fair in May, from May 11 to May 13. A food truck stands in front of building B on those days.",
        "You pay with your staff badge. Guests pay in cash at the till next to the entrance.",
    ]),
}

SIGN_IN_PAGE = "<!DOCTYPE html><html><head><title>Sign In</title></head><body><h1>Sign In</h1><form method=\"post\"><input name=\"username\"><input name=\"password\" type=\"password\"><button>Sign in</button></form></body></html>"


def local_network_address():
    """The address of this computer in the local network, found without sending anything."""
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as probe:
        try:
            probe.connect(("192.0.2.1", 9))
            return probe.getsockname()[0]
        except OSError:
            return None


def render_page(title, paragraphs):
    body = "".join(f"<p>{paragraph}</p>" for paragraph in paragraphs)
    return f"<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>{title}</title></head><body><main><article><h1>{title}</h1>{body}</article></main></body></html>"


class Handler(BaseHTTPRequestHandler):
    server_version = "SharePointMock"

    def do_GET(self):
        url = urlsplit(self.path)
        path = unquote(url.path)

        if path == "/_api/search/query":
            self.answer_search(parse_qs(url.query))
        elif path in PDFS:
            self.answer(200, "application/pdf", (FIXTURES / PDFS[path]).read_bytes())
        elif path in PAGES:
            self.answer(200, "text/html; charset=utf-8", render_page(*PAGES[path]).encode("utf-8"))
        else:
            self.answer(404, "text/plain; charset=utf-8", b"The stand-in knows no such page.")

    def answer_search(self, parameters):
        if self.server.sign_in_page:
            self.answer(200, "text/html; charset=utf-8", SIGN_IN_PAGE.encode("utf-8"))
            return

        query_text = parameters.get("querytext", [""])[0]
        start_row = parameters.get("startrow", ["0"])[0]
        fixture = "sharepoint-search-empty.xml" if "nothing" in query_text.lower() else FIXTURE_BY_START_ROW.get(start_row, "sharepoint-search-empty.xml")

        answer = (FIXTURES / fixture).read_text(encoding="utf-8")
        answer = FIXTURE_HOSTS.sub(self.server.public_address, answer)
        self.answer(200, "application/xml; charset=utf-8", answer.encode("utf-8"))

    def answer(self, status, content_type, body):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    parser = argparse.ArgumentParser(description="A stand-in for a SharePoint Server, see the top of this file.")
    parser.add_argument("--name", default="intranet.test", help="the host name AI Studio is configured with (default: intranet.test)")
    parser.add_argument("--bind", help="the address to listen on (default: the address of this computer in the local network)")
    parser.add_argument("--port", type=int, default=8443, help="the port to listen on (default: 8443)")
    parser.add_argument("--cert", help="the certificate for the host name, as mkcert writes it")
    parser.add_argument("--key", help="the key of the certificate")
    parser.add_argument("--sign-in-page", action="store_true", help="answer every search with a sign-in page")
    parser.add_argument("--print-address", action="store_true", help="print the line for /etc/hosts and stop")
    arguments = parser.parse_args()

    address = arguments.bind or local_network_address()
    if address is None:
        sys.exit("This computer has no address in a local network. Name one with --bind.")

    if arguments.print_address:
        print(f"{address} {arguments.name}")
        return

    if not arguments.cert or not arguments.key:
        sys.exit("AI Studio accepts HTTPS only. Pass --cert and --key, see the top of this file.")

    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(Path(arguments.cert).expanduser(), Path(arguments.key).expanduser())

    server = ThreadingHTTPServer((address, arguments.port), Handler)
    server.socket = context.wrap_socket(server.socket, server_side=True)
    server.public_address = f"https://{arguments.name}:{arguments.port}"
    server.sign_in_page = arguments.sign_in_page

    print(f"Listening on {address}:{arguments.port}. SharePoint Base URL for AI Studio: {server.public_address}/")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print()


if __name__ == "__main__":
    main()
