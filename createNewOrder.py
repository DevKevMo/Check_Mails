import json

with open('order_data.json') as f:
   data = json.load(f)

# Iterate through the list of objects
for obj in data:
    # Find all the orders with "inkl. Kochevent" as the ticket
    kochevent_orders = [order for order in obj["orders"]
                        if order["Ticket"] == "inkl. Kochevent"]

    # Check if there are any "Tagesveranstaltung" orders
    has_tagesveranstaltung = any(
        order["Ticket"] == "Tagesveranstaltung" for order in obj["orders"])

    # If there are both "inkl. Kochevent" and "Tagesveranstaltung" orders, remove the "Tagesveranstaltung" orders
    if kochevent_orders and has_tagesveranstaltung:
        obj["orders"] = [order for order in obj["orders"]
                         if order["Ticket"] != "Tagesveranstaltung"]

with open("newOrder.json", 'w') as json_file:
    json.dump(data, json_file, indent=4)