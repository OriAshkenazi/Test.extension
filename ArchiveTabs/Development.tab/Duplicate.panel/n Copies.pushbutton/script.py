#! python3
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Collections.Generic import List
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import MessageBox, Form, Label, TextBox, Button, DialogResult
from System.Drawing import Point

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def show_message(message, title="Info"):
    """Display a message box"""
    MessageBox.Show(message, title)

def get_user_input(prompt, title, default="2"):
    """Create a simple input dialog using Windows Forms"""
    form = Form()
    form.Text = title
    form.Width = 360
    form.Height = 150
    
    label = Label()
    label.Text = prompt
    label.Location = Point(10, 10)
    label.Width = 350
    
    textbox = TextBox()
    textbox.Text = default
    textbox.Location = Point(10, 40)
    textbox.Width = 260
    
    button = Button()
    button.Text = "OK"
    button.DialogResult = DialogResult.OK
    button.Location = Point(100, 70)
    
    form.Controls.Add(label)
    form.Controls.Add(textbox)
    form.Controls.Add(button)
    form.AcceptButton = button
    
    if form.ShowDialog() == DialogResult.OK:
        return textbox.Text
    return None

def main():
    """
    Creates multiple copies of a selected Revit element.
    Compatible with Revit 2019 and later.
    """
    try:
        # Prompt the user to select an element
        selection = uidoc.Selection
        try:
            ref_picked_obj = selection.PickObject(ObjectType.Element, "Select an element to duplicate")
        except:
            show_message("Selection cancelled by user.")
            return

        picked_element = doc.GetElement(ref_picked_obj.ElementId)
        if not picked_element:
            show_message("Invalid element selected.")
            return

        # Ask for number of copies
        num_copies = get_user_input(
            "Enter the total number of elements (including the original):",
            "Number of Copies"
        )
        
        if not num_copies or not num_copies.isdigit() or int(num_copies) <= 1:
            show_message("Please enter a valid number greater than 1.")
            return
        
        num_copies = int(num_copies)

        # Start transaction
        t = Transaction(doc, "Duplicate Selected Element")
        t.Start()
        
        try:
            # Create copies
            for _ in range(num_copies - 1):
                ElementTransformUtils.CopyElement(doc, picked_element.Id, XYZ(0, 0, 0))
            t.Commit()
            show_message(f"{num_copies} total elements created successfully.")
        except Exception as e:
            t.RollBack()
            print(f"Failed to create copies: {str(e)}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == '__main__':
    main()
