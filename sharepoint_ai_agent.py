import os
from shareplum import Site, Office365
from shareplum.site import Version
from openai import OpenAI

# Configuration
SHAREPOINT_URL = "https://your-sharepoint-site-url"
SHAREPOINT_SITE = "https://your-sharepoint-site-url/sites/your-site-name"
SHAREPOINT_FOLDER = "Shared Documents/Your-Folder-Name"
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SHAREPOINT_USERNAME = os.getenv("SHAREPOINT_USERNAME")
SHAREPOINT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")

# Initialize OpenAI client
openai_client = OpenAI(api_key=OPENAI_API_KEY)

def authenticate_sharepoint(username, password, site_url):
    """Authenticate with SharePoint and return a Site object."""
    authcookie = Office365(site_url, username=username, password=password).GetCookies()
    site = Site(site_url, version=Version.v365, authcookie=authcookie)
    return site

def get_documents_from_sharepoint(site, folder_path):
    """Retrieve documents from a specific SharePoint folder."""
    folder = site.Folder(folder_path)
    return folder.files

def analyze_document_with_ai(document_content):
    """Analyze document content using OpenAI's GPT model."""
    response = openai_client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are a helpful AI assistant that analyzes documents and provides insights."},
            {"role": "user", "content": f"Analyze the following document and provide a summary and key insights:\n\n{document_content}"}
        ]
    )
    return response.choices[0].message.content

def main():
    # Authenticate with SharePoint
    site = authenticate_sharepoint(SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD, SHAREPOINT_SITE)

    # Retrieve documents from the specified folder
    documents = get_documents_from_sharepoint(site, SHAREPOINT_FOLDER)

    # Analyze each document
    for doc in documents:
        print(f"Analyzing document: {doc['Name']}")
        document_content = doc.get_content().decode("utf-8")  # Assuming text-based documents
        analysis_result = analyze_document_with_ai(document_content)
        print(f"Analysis Result:\n{analysis_result}\n")

if __name__ == "__main__":
    main()
