# File to run the quote generation system

import json
from math import prod
import os

import tkinter as tk
from tkinter import Entry, Label, ttk, Button, Listbox, Text, messagebox, END

import process_spreadsheet
from create_template import (
    template_head,
    create_category_a,
    create_category_b,
    create_category_c,
    final_totals,
)


curr_dir = os.getcwd()
category_a_quote = []
category_b_quote = []
category_c_quote = []


# Function to clear all variable values


def clear_variables(
    client_name,
    contact_info,
    requested_by,
    area_on_site,
    work,
    markup_a,
    markup_b,
    markup_c,
    notes,
    category_a_quote,
    category_b_quote,
    category_c_quote,
    listbox_a,
    listbox_b,
    listbox_c,
    combobox_a,
    combobox_b,
    combobox_c,
    quote_name,
):

    client_name.delete(0, "end")
    contact_info.delete(0, "end")
    requested_by.delete(0, "end")
    area_on_site.delete(0, "end")
    work.delete(0, "end")

    markup_a.delete(0, "end")
    markup_b.delete(0, "end")
    markup_c.delete(0, "end")

    notes.delete("1.0", "end")

    category_a_quote.clear()
    category_b_quote.clear()
    category_c_quote.clear()

    combobox_a.delete(0, END)
    combobox_b.delete(0, END)
    combobox_c.delete(0, END)

    listbox_a.delete(0, END)
    listbox_b.delete(0, END)
    listbox_c.delete(0, END)

    quote_name.delete(0, "end")


# Function to merge lists with identical ID's


def merge_duplicates(items):
    # Check if item already exists in list and if so just add quantity to it
    dictionary = {}
    for x in items:
        if x[0] in dictionary.keys():
            dictionary[x[0]] = (x[1], dictionary[x[0]][1] + x[2], x[3])
        else:
            dictionary[x[0]] = (x[1], x[2], x[3])
    ans = []
    for k, v in dictionary.items():
        ans.append([k, v[0], v[1], v[2]])

    return ans


# Function to check if mark up is valid


def is_float(number) -> bool:
    try:
        float(number)
        return True
    except ValueError:
        return False


# Function to validate header text fields


def validate_header(
    client_name,
    contact_info,
    requested_by,
    area_on_site,
    work,
    markup_a,
    markup_b,
    markup_c,
    quote_name,
):

    if client_name == "":
        messagebox.showerror("Form Error", "Please enter a name for the client")
        return False

    elif quote_name == "":
        messagebox.showerror("Form Error", "Please enter a quote name")
        return False

    elif contact_info == "":
        messagebox.showerror("Form Error", "Please enter a contact info for the client")
        return False

    elif requested_by == "":
        messagebox.showerror("Form Error", "Please enter a who requested quote")
        return False

    elif area_on_site == "":
        messagebox.showerror("Form Error", "Please enter the area on site")
        return False

    elif work == "":
        messagebox.showerror("Form Error", "Please enter the work done")
        return False

    elif is_float(markup_a) == False:
        messagebox.showerror(
            "Form Error", "Please enter a mark up A value (without the percentage sign)"
        )
        return False

    elif is_float(markup_b) == False:
        messagebox.showerror(
            "Form Error", "Please enter a mark up B value (without the percentage sign)"
        )
        return False

    elif is_float(markup_c) == False:
        messagebox.showerror(
            "Form Error", "Please enter a mark up C value (without the percentage sign)"
        )
        return False

    else:
        return True


# Function to clear all variable values


def validate_categories(quoted_itemsA, quoted_itemsB, quoted_itemsC):

    if not quoted_itemsA:
        messagebox.showerror("Form Error", "Please enter items in Category A")
        return False

    elif not quoted_itemsB:
        messagebox.showerror("Form Error", "Please enter items for Category B")
        return False

    elif not quoted_itemsC:
        messagebox.showerror("Form Error", "Please enter items for Category C")
        return False

    else:
        return True


# Function to insert value into listBox A


def insert_into_listbox_a(items, quantity_a):
    list_box_a.delete(0, END)

    for item in items:
        item_row = str(item[0]) + ": " + item[1] + "   " + str(item[2])
        list_box_a.insert(END, item_row)

    quantity_a.delete(0, END)


# Function to insert value into listBox B


def insert_into_listbox_b(items, quantity_b):
    list_box_b.delete(0, END)

    for item in items:
        item_row = str(item[0]) + ": " + item[1] + "   " + str(item[2])
        list_box_b.insert(END, item_row)

    quantity_b.delete(0, END)


# Function to insert value into listBox C


def insert_into_listbox_c(items, quantity_c):
    list_box_c.delete(0, END)

    for item in items:
        item_row = str(item[0]) + ": " + item[1] + "   " + str(item[2])
        list_box_c.insert(END, item_row)

    quantity_c.delete(0, END)


# Function to remove value from listBox A


def remove_from_listbox_a(category_a_quote):
    for i in list_box_a.curselection():
        list_box_a.delete(i)
        category_a_quote.pop(i)


# Function to remove value from listBox B


def remove_from_listbox_b(category_b_quote):
    for i in list_box_b.curselection():
        list_box_b.delete(i)
        category_b_quote.pop(i)


# Function to remove value from listBox C


def remove_from_listbox_c(category_c_quote):
    for i in list_box_c.curselection():
        list_box_c.delete(i)
        category_c_quote.pop(i)


# Function to add items from Category A to list


def addToCategoryA(item, quantity_a):
    global category_a_quote

    description = item["Description"]
    quantity = item["Quantity"]
    product_id = description.split(':')[0]

    if quantity == "":
        messagebox.showerror("Invalid Entry", "Please enter a valid quantity")

    else:
        list_data_categoryA = curr_dir + "/json/categoryA.json"
        fob = open(
            list_data_categoryA,
        )
        data = json.load(fob)
        for items in data:
            if product_id == items[0]:
                print(items)
                items = [int(items[0]), items[1], float(quantity), float(items[2])]

                category_a_quote.append(items)

        category_a_quote = merge_duplicates(category_a_quote)
        insert_into_listbox_a(category_a_quote, quantity_a)


# Function to add items from Category B to list


def addToCategoryB(item, quantity_b):
    global category_b_quote

    description = item["Description"]
    quantity = item["Quantity"]

    product_id = description.split(':')[0]

    if quantity == "":
        messagebox.showerror("Invalid Entry", "Please enter a valid quantity")

    else:
        list_data_categoryB = curr_dir + "/json/categoryB.json"
        fob = open(
            list_data_categoryB,
        )
        data = json.load(fob)

        for items in data:
            if product_id == items[0]:
                items = [int(items[0]), items[1], float(quantity), items[2]]
                category_b_quote.append(items)

        category_b_quote = merge_duplicates(category_b_quote)
        insert_into_listbox_b(category_b_quote, quantity_b)


# Function to add items from Category C to list


def addToCategoryC(item, quantity_c):
    global category_c_quote

    description = item["Description"]
    quantity = item["Quantity"]

    product_id = description.split(':')[0]

    if quantity == "":
        messagebox.showerror("Invalid Entry", "Please enter a valid quantity")

    else:
        list_data_categoryC = curr_dir + "/json/categoryC.json"
        fob = open(
            list_data_categoryC,
        )
        data = json.load(fob)

        for items in data:
            if product_id == items[0]:
                if items[3] == "Yes":
                    items = [
                        int(items[0]),
                        items[1] + " (Outsourced)",
                        float(quantity),
                        items[2],
                    ]
                else:
                    items = [int(items[0]), items[1], float(quantity), items[2]]

                category_c_quote.append(items)

        category_c_quote = merge_duplicates(category_c_quote)
        insert_into_listbox_c(category_c_quote, quantity_c)


# Function that gets items for Category A


def getCategoryA():
    list_data_categoryA = curr_dir + "/json/categoryA.json"
    fob = open(
        list_data_categoryA,
    )
    data = json.load(fob)
    categoryA_items = []

    for items in data:
        categoryA_items.append(items[0] + ": " + items[1])

    return categoryA_items


# Function that gets items for Category B


def getCategoryB():
    list_data_categoryB = curr_dir + "/json/categoryB.json"
    fob = open(
        list_data_categoryB,
    )
    data = json.load(fob)
    categoryB_items = []

    for items in data:
        categoryB_items.append(items[0] + ": " + items[1])

    return categoryB_items


# Function that gets items for Category C


def getCategoryC():
    list_data_categoryC = curr_dir + "/json/categoryC.json"
    fob = open(
        list_data_categoryC,
    )
    data = json.load(fob)
    categoryC_items = []

    for items in data:
        categoryC_items.append(items[0] + ": " + items[1])

    return categoryC_items


def create_quote(
    client_name,
    contact_info,
    requested_by,
    area_on_site,
    work,
    categoryA_markup,
    categoryB_markup,
    categoryC_markup,
    notes,
    category_a_quote,
    category_b_quote,
    category_c_quote,
    list_box_a,
    list_box_b,
    list_box_c,
    combobox_a,
    combobox_b,
    combobox_c,
    quote_name,
):
    if (
        validate_header(
            client_name.get(),
            contact_info.get(),
            requested_by.get(),
            area_on_site.get(),
            work.get(),
            categoryA_markup.get(),
            categoryB_markup.get(),
            categoryC_markup.get(),
            quote_name.get(),
        )
        == True
    ):
        if (
            validate_categories(category_a_quote, category_b_quote, category_c_quote)
            == True
        ):
            template_head(
                client_name.get(),
                contact_info.get(),
                requested_by.get(),
                area_on_site.get(),
                work.get(),
            ),
            create_category_a(category_a_quote, categoryA_markup.get() + "%"),
            create_category_b(category_b_quote, categoryB_markup.get() + "%"),
            create_category_c(category_c_quote, categoryC_markup.get() + "%"),
            final_totals(quote_name.get(), notes.get("1.0", "end-1c")),
            clear_variables(
                client_name,
                contact_info,
                requested_by,
                area_on_site,
                work,
                categoryA_markup,
                categoryB_markup,
                categoryC_markup,
                notes,
                category_a_quote,
                category_b_quote,
                category_c_quote,
                list_box_a,
                list_box_b,
                list_box_c,
                combobox_a,
                combobox_b,
                combobox_c,
                quote_name,
            ),
            messagebox.showinfo("Quote Created", "Quote Successfully Generated")


# Function that generates the GUI


def createGUI():

    window = tk.Tk()

    # Get basic information

    heading = Label(
        window, text="Keyan's Invoice Automaton: ", fg="black", font=("Helvetica", 12)
    ).grid(row=0, column=1)
    _clientName = Label(
        window, text="Client Name: ", fg="black", font=("Helvetica", 12)
    ).grid(row=1, column=0)
    client_name = Entry(window, width=30)
    client_name.grid(row=1, column=1)
    _quote_name = Label(
        window, text="Quote Name: ", fg="black", font=("Helvetica", 12)
    ).grid(row=1, column=2)
    quote_name = Entry(window)
    quote_name.grid(row=1, column=3)

    _ContactInfo = Label(
        window, text="Contact Info: ", fg="black", font=("Helvetica", 12)
    ).grid(row=2, column=0)
    contact_info = Entry(window, width=30)
    contact_info.grid(row=2, column=1)
    _markup_categoryA = Label(
        window, text="Category A Mark Up: ", fg="black", font=("Helvetica", 12)
    ).grid(row=2, column=2)
    categoryA_markup = Entry(window)
    categoryA_markup.grid(row=2, column=3)

    _requestedBy = Label(
        window, text="Requested By: ", fg="black", font=("Helvetica", 12)
    ).grid(row=3, column=0)
    requested_by = Entry(window, width=30)
    requested_by.grid(row=3, column=1)
    _markup_categoryB = Label(
        window, text="Category B Mark Up: ", fg="black", font=("Helvetica", 12)
    ).grid(row=3, column=2)
    categoryB_markup = Entry(
        window,
    )
    categoryB_markup.grid(row=3, column=3)

    _areaOnSite = Label(
        window, text="Area on site: ", fg="black", font=("Helvetica", 12)
    ).grid(row=4, column=0)
    area_on_site = Entry(window, width=30)
    area_on_site.grid(row=4, column=1)
    _markup_categoryC = Label(
        window, text="Category C Mark Up: ", fg="black", font=("Helvetica", 12)
    ).grid(row=4, column=2)
    categoryC_markup = Entry(
        window,
    )
    categoryC_markup.grid(row=4, column=3)

    _work = Label(window, text="Work: ", fg="black", font=("Helvetica", 12)).grid(
        row=5, column=0
    )
    work = Entry(window, width=80)
    work.grid(row=5, column=1, columnspan=3)

    # Create Category A section

    categoryA_items = getCategoryA()
    global list_box_a

    _categoryA = Label(window, text="Category A", font=("Helvetica", 14)).grid(
        row=6, column=1, columnspan=2
    )
    _descriptionA = Label(window, text=" Description", font=("Helvetica", 12)).grid(
        row=7, column=1, columnspan=2, sticky=tk.W
    )
    _quantityA = Label(window, text="Quantity", font=("Helvetica", 12)).grid(
        row=7, column=2, columnspan=2, sticky=tk.W
    )
    cb1 = ttk.Combobox(window, values=categoryA_items)
    cb1.grid(row=8, column=1, padx=10, pady=20)

    categoryAQuantity = Entry(window)
    categoryAQuantity.grid(row=8, column=2)

    btnAddCategoryA = Button(
        window,
        text="Add Item",
        fg="blue",
        command=(
            lambda: [
                addToCategoryA(
                    {"Description": cb1.get(), "Quantity": categoryAQuantity.get()},
                    categoryAQuantity,
                ),
            ]
        ),
    )
    btnAddCategoryA.grid(row=8, column=3)

    list_box_a = Listbox(window, width=80, height=5, selectmode="single")
    list_box_a.grid(row=9, column=0, columnspan=3)

    btn_remove_from_listbox_a = Button(
        window,
        text="Remove Item",
        fg="blue",
        command=(lambda: [remove_from_listbox_a(category_a_quote)]),
    )
    btn_remove_from_listbox_a.grid(row=9, column=3)

    # Create Category B section

    categoryB_items = getCategoryB()
    global list_box_b

    _categoryB = Label(window, text="Category B", font=("Helvetica", 14)).grid(
        row=10, column=1, columnspan=2
    )
    _descriptionB = Label(window, text=" Description", font=("Helvetica", 12)).grid(
        row=11, column=1, columnspan=2, sticky=tk.W
    )
    _quantityB = Label(window, text="Quantity", font=("Helvetica", 12)).grid(
        row=11, column=2, columnspan=2, sticky=tk.W
    )
    cb2 = ttk.Combobox(window, values=categoryB_items)
    cb2.grid(row=12, column=1, padx=11, pady=20)

    categoryBQuantity = Entry(window)
    categoryBQuantity.grid(row=12, column=2)

    btnAddCategoryB = Button(
        window,
        text="Add Item",
        fg="blue",
        command=(
            lambda: addToCategoryB(
                {"Description": cb2.get(), "Quantity": categoryBQuantity.get()},
                categoryBQuantity,
            )
        ),
    )
    btnAddCategoryB.grid(row=12, column=3)

    list_box_b = Listbox(window, width=80, height=5, selectmode="single")
    list_box_b.grid(row=13, column=0, columnspan=3)

    btn_remove_from_listbox_b = Button(
        window,
        text="Remove Item",
        fg="blue",
        command=(lambda: remove_from_listbox_b(category_b_quote)),
    )
    btn_remove_from_listbox_b.grid(row=13, column=3)

    # Create Category C section

    categoryC_items = getCategoryC()
    global list_box_c

    _categoryC = Label(window, text="Category C", font=("Helvetica", 14)).grid(
        row=14, column=1, columnspan=2
    )
    _descriptionC = Label(window, text=" Description", font=("Helvetica", 12)).grid(
        row=15, column=1, columnspan=2, sticky=tk.W
    )
    _quantityC = Label(window, text="Quantity", font=("Helvetica", 12)).grid(
        row=15, column=2, columnspan=2, sticky=tk.W
    )
    cb3 = ttk.Combobox(window, values=categoryC_items)
    cb3.grid(row=16, column=1, padx=10, pady=20)

    categoryCQuantity = Entry(window)
    categoryCQuantity.grid(row=16, column=2)

    btnAddCategoryC = Button(
        window,
        text="Add Item",
        fg="blue",
        command=(
            lambda: addToCategoryC(
                {"Description": cb3.get(), "Quantity": categoryCQuantity.get()},
                categoryCQuantity,
            )
        ),
    )
    btnAddCategoryC.grid(row=16, column=3)

    list_box_c = Listbox(window, width=80, height=5, selectmode="single")
    list_box_c.grid(row=17, column=0, columnspan=3)

    btn_remove_from_listbox_c = Button(
        window,
        text="Remove Item",
        fg="blue",
        command=(lambda: remove_from_listbox_c(category_c_quote)),
    )
    btn_remove_from_listbox_c.grid(row=17, column=3)

    global notes
    _notes = Label(window, text="Notes", font=("Helvetica", 12)).grid(
        row=18, column=2, columnspan=2, sticky=tk.W
    )
    notes = Text(window, width=75, height=10, bg="light yellow")
    notes.grid(row=19, column=0, columnspan=4)

    btn3 = Button(
        window,
        text="Generate Quote",
        fg="blue",
        command=(
            lambda: [
                create_quote(
                    client_name=client_name,
                    contact_info=contact_info,
                    requested_by=requested_by,
                    area_on_site=area_on_site,
                    work=work,
                    categoryA_markup=categoryA_markup,
                    categoryB_markup=categoryB_markup,
                    categoryC_markup=categoryC_markup,
                    notes=notes,
                    category_a_quote=category_a_quote,
                    category_b_quote=category_b_quote,
                    category_c_quote=category_c_quote,
                    list_box_a=list_box_a,
                    list_box_b=list_box_b,
                    list_box_c=list_box_c,
                    combobox_a=cb1,
                    combobox_b=cb2,
                    combobox_c=cb3,
                    quote_name=quote_name,
                ),
            ]
        ),
    )
    btn3.grid(row=20, column=1, columnspan=2)

    window.title("Keyan's Smart Invoice Generator")
    window.geometry("620x1080")
    window.mainloop()


# Function that is the base of the application


def main():

    process_spreadsheet.main()
    createGUI()


if __name__ == "__main__":
    main()
