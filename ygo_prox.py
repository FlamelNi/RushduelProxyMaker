import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches
from io import BytesIO

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

def get_card_image(card_name):
    """
    Search DeviantArt for 'rush duel <card_name>',
    follow the first result, and return the full-size image URL.
    """
    query = f"rush duel {card_name}".replace(" ", "+")
    search_url = f"https://www.deviantart.com/search?q={query}"

    try:
        resp = requests.get(search_url, headers=headers, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        # Find first artwork link
        first_link = soup.select_one("a[href][aria-label][aria-label*='by']")
        if not first_link:
            return None

        art_url = first_link["href"]

        # Step 2: Artwork page
        art_resp = requests.get(art_url, headers=headers, timeout=10)
        art_resp.raise_for_status()
        art_soup = BeautifulSoup(art_resp.text, "html.parser")

        # Find full-size image
        img_tag = art_soup.select_one("div[typeof=ImageObject] img")
        if img_tag:
            return img_tag["src"]

    except Exception as e:
        print(f"Error fetching {card_name}: {e}")

    return None


def decklist_to_docx(deck_file, output_docx, card_width=2.31, card_height=3.37, per_row=3):
    """
    Reads a decklist, fetches card images from DeviantArt, and builds a Word document.
    """
    doc = Document()

    # Set margins (like before)
    section = doc.sections[0]
    from docx.shared import Inches
    section.top_margin = Inches(0.4)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    card_entries = []

    # Parse decklist
    with open(deck_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.lower() in ["monster", "spell", "trap", "extra", "side", "몬스터", "마법", "함정", "엑스트라", "사이드"]:
                continue  # skip headers

            parts = line.split(" ", 1)
            if len(parts) != 2:
                continue
            count, name = parts
            if not count.isdigit():
                continue
            card_entries.append((int(count), name))

    # Build document with images
    for i in range(0, len(card_entries)):
        count, card_name = card_entries[i]
        print(f"Searching: {card_name}")
        img_url = get_card_image(card_name)

        if not img_url:
            print(f"❌ No image found for {card_name}")
            continue

        try:
            img_resp = requests.get(img_url, headers=headers, timeout=10)
            img_resp.raise_for_status()
            img_bytes = BytesIO(img_resp.content)

            # Add cards in rows
            for _ in range(count):
                if doc.paragraphs and len(doc.paragraphs[-1].runs) < per_row:
                    # continue current row
                    run = doc.paragraphs[-1].add_run()
                else:
                    # start new row
                    run = doc.add_paragraph().add_run()

                run.add_picture(img_bytes, width=Inches(card_width), height=Inches(card_height))

        except Exception as e:
            print(f"Error adding {card_name}: {e}")

    doc.save(output_docx)
    print(f"Document saved as {output_docx}")


# 🔎 Example usage
if __name__ == "__main__":
    decklist_to_docx("deck.txt", "deck_proxies.docx")
