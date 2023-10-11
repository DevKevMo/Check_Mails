import pandas as pd
import json

# Your data object (example)
with open('order_data.json') as f:
   data = json.load(f)

df = pd.DataFrame(columns=["OrderNr", "Ticket", "Name", "Company"])

# Iterieren Sie über die Daten und fügen Sie sie dem DataFrame hinzu
for item in data:
    orderNr = item["orderNr"]
    for order in item["orders"]:
        name = order["Name"]
        ticket = order["Ticket"]
        # In diesem Fall bleibt die Company leer
        company = ""
        df = df._append({"OrderNr": orderNr, "Ticket": ticket, "Name": name, "Company": company}, ignore_index=True)

# Speichern Sie den DataFrame in einer Excel-Datei
df.to_excel("output.xlsx", index=False)