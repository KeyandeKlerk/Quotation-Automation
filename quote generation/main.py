# File to run the quote generation system

from asyncio import windows_events
import json
import os

import tkinter as tk
from tkinter import Entry, Label, ttk, Button, Listbox, Text, END
from turtle import width

from create_template import *

curr_dir = os.getcwd()
CategoryAQuote = []
CategoryBQuote = []
CategoryCQuote = []

# Function to insert value into listBox A


def insert_into_listbox_a(items):
    list_box_a.delete(0, END)
    for item in items:
        item_row = item[0] + item[1]
        list_box_a.insert(END, item_row)

# Function to insert value into listBox B


def insert_into_listbox_b(items):
    list_box_b.delete(0, END)
    for item in items:
        item_row = item[0] + item[1]
        list_box_b.insert(END, item_row)

# Function to insert value into listBox C


def insert_into_listbox_c(items):
    list_box_c.delete(0, END)
    for item in items:
        item_row = item[0] + item[1]
        list_box_c.insert(END, item_row)

# Function to remove value from listBox A


def remove_from_listbox_a():
    for i in list_box_a.curselection():
        list_box_a.delete(i)
        CategoryAQuote.pop(i)

# Function to remove value from listBox B


def remove_from_listbox_b():
    for i in list_box_b.curselection():
        list_box_b.delete(i)
        CategoryBQuote.pop(i)

# Function to remove value from listBox C


def remove_from_listbox_c():
    for i in list_box_c.curselection():
        list_box_c.delete(i)
        CategoryCQuote.pop(i)

# Function to add items from Category A to list


def addToCategoryA(item):
    global jsonA
    product_id = item['Description'][0]
    quanitity = item['Quantity']

    json_data_categoryA = curr_dir + '\\json\\categoryA.json'
    fob = open(json_data_categoryA,)
    data = json.load(fob)

    for items in data:
        if product_id == items[0]:

            items = [items[1], quanitity, items[2]]
            CategoryAQuote.append(items)

    jsonA = json.dumps(CategoryAQuote)
    insert_into_listbox_a(CategoryAQuote)


# Function to add items from Category B to list


def addToCategoryB(item):
    global jsonB
    product_id = item['Description'][0]
    quanitity = item['Quantity']

    json_data_categoryB = curr_dir + '\\json\\categoryB.json'
    fob = open(json_data_categoryB,)
    data = json.load(fob)

    for items in data:
        if product_id == items[0]:
            items = [items[1], quanitity, items[2]]
            CategoryBQuote.append(items)

    jsonB = json.dumps(CategoryBQuote)
    insert_into_listbox_b(CategoryBQuote)

# Function to add items from Category C to list


def addToCategoryC(item):
    global jsonC
    product_id = item['Description'][0]
    quanitity = item['Quantity']

    json_data_categoryC = curr_dir + '\\json\\categoryC.json'
    fob = open(json_data_categoryC,)
    data = json.load(fob)

    for items in data:
        print(items)
        if product_id == items[0]:
            if items[3] == 'Yes':
                items = [''.join((items[1] + ' (Outsourced)')),
                         quanitity, items[2], items[3]]
            else:
                items = [items[1], quanitity, items[2], items[3]]

            CategoryCQuote.append(items)

    jsonC = json.dumps(CategoryCQuote)
    insert_into_listbox_c(CategoryCQuote)

# Function that gets items for Category A


def getCategoryA():
    json_data_categoryA = curr_dir + '\\json\\categoryA.json'
    fob = open(json_data_categoryA,)
    data = json.load(fob)
    categoryA_items = []

    for items in data:
        categoryA_items.append(items[0] + ': ' + items[1])

    return categoryA_items

# Function that gets items for Category B


def getCategoryB():
    json_data_categoryB = curr_dir + '\\json\\categoryB.json'
    fob = open(json_data_categoryB,)
    data = json.load(fob)
    categoryB_items = []

    for items in data:
        categoryB_items.append(items[0] + ': ' + items[1])

    return categoryB_items

# Function that gets items for Category C


def getCategoryC():
    json_data_categoryC = curr_dir + '\\json\\categoryC.json'
    fob = open(json_data_categoryC,)
    data = json.load(fob)
    categoryC_items = []

    for items in data:
        categoryC_items.append(items[0] + ': ' + items[1])

    return categoryC_items

# Function that generates the GUI


def createGUI():

    window = tk.Tk()

    # Get basic information

    heading = Label(window, text="Keyan's Invoice Automaton: ",
                    fg='black', font=("Helvetica", 12)).grid(row=0, column=1)
    _clientName = Label(window, text="Client Name: ",
                        fg='black', font=("Helvetica", 12)).grid(row=1, column=0)
    client_name = Entry(window,  width=30)
    client_name.grid(row=1, column=1)
    _markup_categoryA = Label(window, text="Category A Mark Up: ",
                              fg='black', font=("Helvetica", 12)).grid(row=1, column=2)
    categoryA_markup = Entry(window)
    categoryA_markup.grid(row=1, column=3)

    _ContactInfo = Label(window, text="Contact Info: ",
                         fg='black', font=("Helvetica", 12)).grid(row=2, column=0)
    contact_info = Entry(window,  width=30)
    contact_info.grid(row=2, column=1)
    _markup_categoryB = Label(window, text="Category B Mark Up: ",
                              fg='black', font=("Helvetica", 12)).grid(row=2, column=2)
    categoryB_markup = Entry(window)
    categoryB_markup.grid(row=2, column=3)

    _requestedBy = Label(window, text="Requested By: ",
                         fg='black', font=("Helvetica", 12)).grid(row=3, column=0)
    requested_by = Entry(window,  width=30)
    requested_by.grid(row=3, column=1)
    _markup_categoryC = Label(window, text="Category C Mark Up: ",
                              fg='black', font=("Helvetica", 12)).grid(row=3, column=2)
    categoryC_markup = Entry(window,)
    categoryC_markup.grid(row=3, column=3)

    _areaOnSite = Label(window, text="Area on site: ",
                        fg='black', font=("Helvetica", 12)).grid(row=4, column=0)
    area_on_site = Entry(window,  width=30)
    area_on_site.grid(row=4, column=1)

    _work = Label(window, text="Work: ",
                  fg='black', font=("Helvetica", 12)).grid(row=5, column=0)
    work = Entry(window, width=80)
    work.grid(row=5, column=1, columnspan=3)

    # Create Category A section

    categoryA_items = getCategoryA()
    global list_box_a

    _categoryA = Label(window, text="Category A", font=("Helvetica", 14)
                       ).grid(row=6, column=1, columnspan=2)
    _descriptionA = Label(window, text=" Description", font=("Helvetica", 12)
                          ).grid(row=7, column=1, columnspan=2, sticky=tk.W)
    _quantityA = Label(window, text="Quantity", font=("Helvetica", 12)
                       ).grid(row=7, column=2, columnspan=2, sticky=tk.W)
    cb1 = ttk.Combobox(window, values=categoryA_items)
    cb1.grid(row=8, column=1, padx=10, pady=20)

    categoryAQuantity = Entry(window)
    categoryAQuantity.grid(row=8, column=2)

    btnAddCategoryA = Button(window, text="Add Item",
                             fg='blue', command=(lambda: addToCategoryA({"Description": cb1.get(), "Quantity": categoryAQuantity.get()})))
    btnAddCategoryA.grid(row=8, column=3)

    list_box_a = Listbox(
        window, width=80, height=5, selectmode='single')
    list_box_a.grid(row=9, column=0, columnspan=3)

    btn_remove_from_listbox_a = Button(window, text="Remove Item",
                                       fg='blue', command=(lambda: remove_from_listbox_a()))
    btn_remove_from_listbox_a.grid(row=9, column=3)

    # Create Category B section

    categoryB_items = getCategoryB()
    global list_box_b

    _categoryB = Label(window, text="Category B", font=("Helvetica", 14)
                       ).grid(row=10, column=1, columnspan=2)
    _descriptionB = Label(window, text=" Description", font=("Helvetica", 12)
                          ).grid(row=11, column=1, columnspan=2, sticky=tk.W)
    _quantityB = Label(window, text="Quantity", font=("Helvetica", 12)
                       ).grid(row=11, column=2, columnspan=2, sticky=tk.W)
    cb2 = ttk.Combobox(window, values=categoryB_items)
    cb2.grid(row=12, column=1, padx=11, pady=20)

    categoryBQuantity = Entry(window)
    categoryBQuantity.grid(row=12, column=2)

    btnAddCategoryB = Button(window, text="Add Item",
                             fg='blue', command=(lambda: addToCategoryB({"Description": cb2.get(), "Quantity": categoryBQuantity.get()})))
    btnAddCategoryB.grid(row=12, column=3)

    list_box_b = Listbox(window, width=80, height=5, selectmode='single')
    list_box_b.grid(row=13, column=0, columnspan=3)

    btn_remove_from_listbox_b = Button(window, text="Remove Item",
                                       fg='blue', command=(lambda: remove_from_listbox_b()))
    btn_remove_from_listbox_b.grid(row=13, column=3)

    # Create Category C section

    categoryC_items = getCategoryC()
    global list_box_c

    _categoryC = Label(window, text="Category C", font=("Helvetica", 14)
                       ).grid(row=14, column=1, columnspan=2)
    _descriptionC = Label(window, text=" Description", font=("Helvetica", 12)
                          ).grid(row=15, column=1, columnspan=2, sticky=tk.W)
    _quantityC = Label(window, text="Quantity", font=("Helvetica", 12)
                       ).grid(row=15, column=2, columnspan=2, sticky=tk.W)
    cb3 = ttk.Combobox(window, values=categoryC_items)
    cb3.grid(row=16, column=1, padx=10, pady=20)

    categoryCQuantity = Entry(window)
    categoryCQuantity.grid(row=16, column=2)

    btnAddCategoryC = Button(window, text="Add Item",
                             fg='blue', command=(lambda: addToCategoryC({"Description": cb3.get(), "Quantity": categoryCQuantity.get()})))
    btnAddCategoryC.grid(row=16, column=3)

    list_box_c = Listbox(window, width=80, height=5, selectmode='single')
    list_box_c.grid(row=17, column=0, columnspan=3)

    btn_remove_from_listbox_c = Button(window, text="Remove Item",
                                       fg='blue', command=(lambda: remove_from_listbox_c()))
    btn_remove_from_listbox_c.grid(row=17, column=3)

    _notes = Label(window, text="Notes", font=("Helvetica", 12)
                   ).grid(row=18, column=2, columnspan=2, sticky=tk.W)
    notes = Text(window, width=75, height=10, bg="light yellow")
    notes.grid(row=19, column=0, columnspan=4)

    btn3 = Button(window, text="Generate Quote",
                  fg='blue', command=(lambda: [template_head(client_name=client_name.get(), contact_info=contact_info.get(), requested_by=requested_by.get(), area_on_site=area_on_site.get(), work=work.get()), create_category_a(jsonA, categoryA_markup.get() + '%'), create_category_b(jsonB, categoryB_markup.get() + '%'), create_category_c(jsonC, categoryC_markup.get() + '%'), final_totals(notes.get("1.0", 'end-1c'))]))
    btn3.grid(row=20, column=1, columnspan=2)

    window.title('Keyan\'s Smart Invoice Generator')
    window.geometry('620x1080')
    window.mainloop()


# Function that is the base of the application


def main():

    init()
    createGUI()


if __name__ == '__main__':
    main()
