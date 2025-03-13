import argparse
import csv
import logging
import requests
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional
from openpyxl import Workbook

BASE_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
FETCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"

# Fetch papers based on search query
def fetch_papers(query: str) -> List[Dict[str, str]]:
    """Fetch paper IDs from PubMed based on search query."""
    params = {
        "db": "pubmed",
        "term": query,
        "retmode": "json",
        "retmax": 50
    }
    response = requests.get(BASE_URL, params=params)
    response.raise_for_status()
    ids = response.json().get("esearchresult", {}).get("idlist", [])

    papers = []
    for paper_id in ids:
        paper = fetch_paper_details(paper_id)
        if paper:
            papers.append(paper)

    return papers

# Identify company affiliations
def is_company_affiliation(affiliation: str) -> bool:
    """Check if the affiliation is from a company based on keywords."""
    keywords = [
        "Pharma", "Inc.", "Ltd.", "Company", "Biotech", "Corporation", "LLC",
        "Laboratory", "Research", "Institute", "GmbH", "Pvt.", "Technologies",
        "Diagnostics", "Therapeutics", "Healthcare", "Medical", "Pharmaceuticals",
        "Solutions", "Enterprises", "Global", "Holdings", "Industries", "Manufacturing"
    ]
    return any(keyword.lower() in affiliation.lower() for keyword in keywords)

# Fetch details for individual papers
def fetch_paper_details(paper_id: str) -> Optional[Dict[str, str]]:
    """Fetch detailed information for a specific paper."""
    params = {
        "db": "pubmed",
        "id": paper_id,
        "retmode": "xml"
    }
    response = requests.get(FETCH_URL, params=params)
    if response.status_code == 200:
        root = ET.fromstring(response.content)

        title = root.find(".//ArticleTitle")
        title = title.text if title is not None else "N/A"

        pub_date = root.find(".//ArticleDate") if root.find(".//ArticleDate") is not None else root.find(".//PubDate")
        if pub_date is not None:
            year = pub_date.find("Year").text if pub_date.find("Year") is not None else "N/A"
            month = pub_date.find("Month").text if pub_date.find("Month") is not None else "N/A"
            day = pub_date.find("Day").text if pub_date.find("Day") is not None else "N/A"
            publication_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        else:
            publication_date = "N/A"

        non_academic_authors = []
        company_affiliations = []
        authors = root.findall(".//Author")
        for author in authors:
            last_name = author.find("LastName")
            fore_name = author.find("ForeName")
            affiliation = author.find("Affiliation")

            if last_name is not None and fore_name is not None:
                author_name = f"{fore_name.text} {last_name.text}"
                if affiliation is not None:
                    affiliation_text = affiliation.text
                    if is_company_affiliation(affiliation_text):
                        non_academic_authors.append(f"{author_name} ({affiliation_text})")
                        company_affiliations.append(affiliation_text)

        email = None
        affiliations = root.findall(".//Affiliation")
        for affiliation in affiliations:
            if affiliation is not None and "@" in affiliation.text:
                email = affiliation.text
                break
        email = email if email else "N/A"

        return {
            "PubmedID": paper_id,
            "Title": title,
            "Publication Date": publication_date,
            "Non-academic Author(s)": ", ".join(non_academic_authors) if non_academic_authors else "N/A",
            "Company Affiliation(s)": ", ".join(company_affiliations) if company_affiliations else "N/A",
            "Corresponding Author Email": email
        }
    return None

# Save to Excel
def save_to_excel(papers: List[Dict[str, str]], filename: str):
    """Save fetched papers to an Excel file."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "PubMed Papers"

    headers = list(papers[0].keys())
    sheet.append(headers)

    for paper in papers:
        sheet.append(list(paper.values()))

    workbook.save(filename)

# Save to CSV
def save_to_csv(papers: List[Dict[str, str]], filename: str):
    """Save fetched papers to a CSV file."""
    with open(filename, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.DictWriter(file, fieldnames=papers[0].keys())
        writer.writeheader()
        writer.writerows(papers)

# Main function
def main():
    """Main function to handle user input and control flow."""
    parser = argparse.ArgumentParser(description="Fetch papers from PubMed")
    parser.add_argument("query", type=str, help="Search query")
    parser.add_argument("-d", "--debug", action="store_true", help="Enable debug mode")
    parser.add_argument("-f", "--file", type=str, help="Output CSV file")
    parser.add_argument("-x", "--excel", type=str, help="Output Excel file")

    args = parser.parse_args()

    if args.debug:
        logging.basicConfig(level=logging.DEBUG)

    papers = fetch_papers(args.query)

    if args.file:
        save_to_csv(papers, args.file)
        print(f"Results saved to {args.file}")

    if args.excel:
        save_to_excel(papers, args.excel)
        print(f"Results saved to {args.excel}")

    if not args.file and not args.excel:
        for paper in papers:
            print("=" * 66)
            print(f"PubMed ID: {paper['PubmedID']}")
            print(f"Title: {paper['Title']}")
            print(f"Publication Date: {paper['Publication Date']}")
            print(f"Non-academic Author(s): {paper['Non-academic Author(s)']}")
            print(f"Company Affiliation(s): {paper['Company Affiliation(s)']}")
            print(f"Corresponding Author Email: {paper['Corresponding Author Email']}")
            print("=" * 66)

if __name__ == "__main__":
    main()
