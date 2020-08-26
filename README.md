# iCiteAnalysisWebScraper
![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)

Python-based webscraper utilizing selenium to extract certain values from the iCiteAnalysis tool using chromium driver. 

## Explanation 

```python
wb = load_workbook("Test Excel.xlsx")
ws = wb["Sheet1"]
column = ws["A"]
names = [column[x].value for x in range(len(column))]
```

In this initialization step, the code takes an Excel file with a list of relevant names within a certain sheet (Sheet1) and column (A). The names are then broken down into a list with values consisting of first and last names. 

```python
rcr_values = []
pub_year = []
avg_hum = []
length = len(names)-1
url = "https://icite.od.nih.gov/analysis"
browser = webdriver.Chrome(r" ")
browser.get(url)
```

This step sets up empty lists for desired values, a max limit of the list used for indexing in the next step, and initializes the chrome driver given the location is inputted within the field.

```python
for x in names:
.
.
.
```

Main portion of the code which cycles through the following steps: 
* Inputs the first name into the search field
  * If no results found, re-searches based on next name
  * If too many results, accepts results and continues
* Determines if the first and last name of the author is present in the authors of publications
  * If not, retries search with the next name
  * If yes, obtains desired values from current page and auxilliary pages

Exceptions account for possible errors in the website based on timeouts and network delays while wait commands serve to lower the possibility of network delay error. 

```python
df = pd.DataFrame()
df["Names"] = names
df["Mean RCRs"] = rcr_values
df["Pubs Per Year"] = pub_year
df["Average Human Score"] = avg_hum
df.to_excel("Excel Results.xlsx", index=False)
```

Creates a new Excel file with the names and desired values in adjacent columns.
