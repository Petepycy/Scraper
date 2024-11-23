import binascii
import requests
from bs4 import BeautifulSoup
import time
import csv
import base64


class PolishScraper:
    def __init__(self, base_url):
        self.base_url = base_url
        self.headers = {"User-Agent": "Mozilla/5.0"}
        self.data = []

    def fetch_page(self, url):
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.text

    def parse_main_page(self):
        print("Parsing main page...")
        html_content = self.fetch_page(self.base_url)
        soup = BeautifulSoup(html_content, "html.parser")
        component_links = []

        components = soup.find_all("li", class_="wizRow")
        for component in components:
            link_tag = component.find("a", class_="wizLnk")
            if link_tag:
                component_links.append(link_tag["href"])

        print(f"Found {len(component_links)} components.")
        return component_links

    def decode_base64(self, encoded_str):
        """Decodes Base64 strings with padding correction."""
        try:
            # Add padding if missing
            missing_padding = len(encoded_str) % 4
            if missing_padding:
                encoded_str += "=" * (4 - missing_padding)
            # Decode Base64
            decoded_str = base64.b64decode(encoded_str).decode("utf-8")
            return decoded_str
        except (binascii.Error, UnicodeDecodeError) as e:
            print(f"Error decoding Base64 string: {encoded_str}")
            print(f"Exception: {e}")
            return None

    def parse_component_page(self, url):
        print(f"Parsing component page: {url}")
        html_content = self.fetch_page(url)
        soup = BeautifulSoup(html_content, "html.parser")

        component_data = {}

        # Address and NIP information
        address_box = soup.find("div", id="addrBox")
        if address_box:
            component_data["Ulica"] = address_box.find("span",
                                                       itemprop="streetAddress").text.strip() if address_box.find(
                "span", itemprop="streetAddress") else None
            component_data["Kod pocztowy"] = address_box.find("span",
                                                              itemprop="postalCode").text.strip() if address_box.find(
                "span", itemprop="postalCode") else None
            component_data["Miasto"] = address_box.find("span",
                                                        itemprop="addressLocality").text.strip() if address_box.find(
                "span", itemprop="addressLocality") else None
            component_data["Województwo"] = address_box.find("span",
                                                             itemprop="addressRegion").text.strip() if address_box.find(
                "span", itemprop="addressRegion") else None

            # Extract NIP
            nip_element = address_box.find("span", text=lambda t: t and "NIP:" in t)
            if nip_element:
                nip_container = nip_element.find_parent("div")  # Find the parent div containing the NIP
                if nip_container:
                    # Get the text of the div and remove the content of the span
                    nip_text = nip_container.get_text(strip=True).replace(nip_element.get_text(strip=True), "").strip()
                    component_data["NIP"] = nip_text

        # Telephones
        phone_box = soup.find("div", id="telBox")
        if phone_box:
            telephones = [tel.text.strip() for tel in phone_box.find_all("span", itemprop="telephone")]
            component_data["Telefony"] = ", ".join(telephones)

        # Branches
        br_box = soup.find("div", id="brBox")
        if br_box:
            branches = [branch.text.strip() for branch in br_box.find_all("span")]
            component_data["Branże"] = ", ".join(branches)

        # Website and Email
        www_box = soup.find("div", id="wwwBox")
        if www_box:
            www_link = www_box.find("a", itemprop="url")
            component_data["Strona WWW"] = www_link["href"] if www_link else None

            email_image = www_box.find("img", class_="emlImg")
            if email_image and "data-src" in email_image.attrs:
                email_encoded = email_image["data-src"].split("?usr=")[-1].split("&dmn=")
                username = self.decode_base64(email_encoded[0]) if len(email_encoded) > 0 else None
                domain = self.decode_base64(email_encoded[1]) if len(email_encoded) > 1 else None
                component_data["Email"] = f"{username}@{domain}" if username and domain else None

        return component_data

    def scrape(self):
        component_links = self.parse_main_page()

        for link in component_links:
            full_url = link if link.startswith("http") else f"https://www.baza-firm.com.pl{link}"
            component_data = self.parse_component_page(full_url)
            self.data.append(component_data)
            time.sleep(1)  # Be polite and avoid overloading the server

    def save_data(self, filename="scraped_data.csv"):
        print(f"Saving data to {filename}...")
        if not self.data:
            print("No data to save.")
            return

        keys = self.data[0].keys()
        with open(filename, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.DictWriter(file, fieldnames=keys)
            writer.writeheader()
            writer.writerows(self.data)
        print("Data saved successfully!")


if __name__ == "__main__":
    scraper = PolishScraper(
        "https://www.baza-firm.com.pl/?vsk=Gdynia+&vn=&vm=&vu=&vw=&vwn=&b_szukaj=szukaj&v_br_opis=&v_br=")
    scraper.scrape()
    scraper.save_data()
