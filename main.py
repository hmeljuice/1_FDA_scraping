import mechanicalsoup
import win32com.client as win32
import os

# kaj iščemo
od_datuma = "01/01/2010"
ime_excel_datoteke = "Podatki iz FDA"
odpri_excel_med_zapisovanjem = True


# link do iskalne strani in odpiranje te strani
stating_link = "https://www.accessdata.fda.gov/scripts/cder/safetylabelingchanges/"

make_browser = mechanicalsoup.Browser()
search_form_link = make_browser.get(stating_link)
search_the_form = search_form_link.soup

# izbira forme in določanje iskalnih parametrov
form = search_the_form.select("form")[2]
form.select("input")[0]["value"] = od_datuma


# submitanje form
results_page = make_browser.submit(form, search_form_link.url)
rezultati = results_page.soup.find_all("tr")

# priprava podatkov v liste

data = []
a = 0
for el in rezultati:
    a += 1

    # prvih 18 je neuporabnih, zato:
    if a > 18:
        b = 0
        for row in el:
            b += 1
            if b == 2:
                # data.append(f"https://www.accessdata.fda.gov{row.a['href']}")
                data.append(row.text.strip())
            elif b == 4:
                data.append(row.text.strip())
            elif b == 6:
                data.append(row.text.strip())
            elif b == 8:
                data.append(row.text.strip())
            elif b == 10:
                data.append(row.text.strip())
            elif b == 12:
                data.append(row.text.strip())



# make excel workbook
excel = win32.Dispatch('Excel.Application')
excel.Visible = odpri_excel_med_zapisovanjem
wb = excel.Workbooks.Add()
wb.Worksheets.Add()
wb.Worksheets[1].Name = "BLA"
wb.Worksheets[2].Name = "NDA"

# write to excel
header_letter = ["A", "B", "C", "D", "E", "F"]
number_of_rows = int(len(data) / 6)
c = 0
e = 0
d = 0

for j in range(1, number_of_rows+1):
    print(j)
    temp = data[d:d+6]
    d += 6

    if temp[3] == "BLA":
        for i in range(1, 7):
            wb.Worksheets(1).Range(f"{header_letter[i-1]}{c+1}").Value = temp[i-1]
        c += 1
    elif temp[3] == "NDA":
        for i in range(1, 7):
            wb.Worksheets(2).Range(f"{header_letter[i-1]}{e+1}").Value = temp[i-1]
        e += 1

# Save workbook at end
wb.SaveAs(os.path.join(os.getcwd(), f"{ime_excel_datoteke}.xlsx"))

