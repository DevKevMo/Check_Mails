import win32com.client
import os
import re
import json

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
subfolder_name = "Kundentag"
directory = os.getcwd()

subfolder = None
for folder in inbox.Folders:
    if folder.Name == subfolder_name:
        subfolder = folder
        break

if subfolder:
    messages = subfolder.Items
    message = messages.GetLast()
    pattern = r"(Bestellung|0,00 €)\s+(.*?)1 x\s+Teilnahme\s+(inkl. Kochevent|Tagesveranstaltung)"
    patternOrderNr = r"Bestellung\s+#\d{10}"
    attachments = message.Attachments

    if attachments.Count > 0:
        barcodeData = []
        for i in range(attachments.Count):
            attachment = attachments.Item(i + 1)

            base_filename, file_extension = os.path.splitext(
                attachment.FileName)
            unique_filename = f"{base_filename}_{i}{file_extension}"

            save_path = os.path.join(directory + "\\mails", unique_filename)
            attachment.SaveAsFile(save_path)
            orders = []
            msg = outlook.OpenSharedItem(save_path)
            matches = re.findall(pattern, msg.body)
            if matches is not None:
                orderNrMatch = re.findall(patternOrderNr, msg.body)
                for match in matches:
                    order = {
                        "Name": match[1].strip(),
                        "Ticket": match[2].strip()
                    }
                    orders.append(order)
                orderData = {
                    "orderNr": orderNrMatch,
                    "orders": orders
                }
                if orderData: 
                    barcodeData.append(orderData)
                
        file_name = 'order_data.json'
        file_path = directory + "\\" + file_name

        with open(file_path, 'w') as json_file:
            json.dump(barcodeData, json_file, indent=4)

        print(f"Data saved to {file_path}")
