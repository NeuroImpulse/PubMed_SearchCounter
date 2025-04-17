import pandas as pd
import requests
import time

# Load your Excel file
df = pd.read_excel("/Users/sunyixiao/Desktop/AD_HippocampalCoding_SearchQueries.xlsx")

# Base PubMed URL
base_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"

# Function to search each query
def search_pubmed_count(query):
    if pd.isna(query) or not isinstance(query, str) or query.strip() == "":
        return 0
    params = {
        "db": "pubmed",
        "term": query,
        "retmode": "json"
    }
    try:
        response = requests.get(base_url, params=params)
        time.sleep(0.3)  # polite delay to PubMed
        if response.status_code == 200:
            data = response.json()
            return int(data["esearchresult"]["count"])
        else:
            print(f"Failed: {query}")
            return None
    except Exception as e:
        print(f"Error with {query}: {e}")
        return None

# Loop with progress
results = []
for i, row in df.iterrows():
    query = row["PubMed_Search_Query"]
    count = search_pubmed_count(query)
    results.append(count)
    print(f"[{i+1}/{len(df)}] {query} → {count} articles")

df["PubMed_Article_Count"] = results

# Save to new file
df.to_excel("AD_HippocampalCoding_SearchQueries_withCounts.xlsx", index=False)
print("✅ Done! Saved with article counts.")