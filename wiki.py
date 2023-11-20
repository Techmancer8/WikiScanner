import json, requests
from tqdm import tqdm
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter

alllist = "https://en.wikipedia.org/wiki/List_of_orbital_launch_systems"

def get_links(link):
    html = requests.get(link)
    page = BeautifulSoup(html.content, "html.parser")

    links = []
    elems = page.find_all("ul")
    for li in elems:
        for a in li.find_all("a"):
            try:
                link = a["href"]
                if link[0] == '#':
                    continue
                if 'https://' not in link:
                    link = 'https://en.wikipedia.org'+link
                links.append(link)
            except KeyError:
                break
    return links

def get_data(link):
    #print(f"searching {link}")
    page = BeautifulSoup(requests.get(link).content, "html.parser")
    infobox = page.find("table",class_="infobox hproduct")
    if infobox is None:
        return None, None
    
    name = page.find("h1",id="firstHeading").text
    #print(name)

    data = dict()
    subset = dict()
    head = None
    for row in infobox.find_all("tr"):
        dat = row.find("td")
        if dat is None:
            if head is not None:
                data[head.text] = subset
                subset = dict()
            head = row.find("th",class_="infobox-header")
        else:
            title = row.find("th")
            if title:
                subset[title.text] = dat.text
        
        #i.find("th",class_="infobox-header",text="Size")
    return name, data

def sheet():
    filename = "IVD/rockets.xlsx"
    workbook = Workbook()
    sheet = workbook.active

    for i in range(2,)
    sheet["A1"] = "hello"
    sheet["B1"] = "world!"

    workbook.save(filename=filename)

def main():
    links = get_links(alllist)
    linkbar = tqdm(links)
    for i in linkbar:
        try:
            name, data = get_data(i)
            
        except KeyboardInterrupt:
            print("\nEnding")
            break
        except Exception as e:
            print(f"\nError: {i}\n{e}")
        finally:
            if name is not None:
                linkbar.set_description(name)

def test():
    json.dump(get_data("https://en.wikipedia.org/wiki/Saturn_V"), open('results.json','w'))

if __name__ == "__main__":
    main()