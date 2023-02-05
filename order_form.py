import dearpygui.dearpygui as dpg
from pdf2image import convert_from_path
import os
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import AnnotationBuilder
import json

class OrderForm:

    def __init__(self):
        self.file_count = 0
        self.orders_path = 'Orders_to_Process/'
        self.completed_orders_path = 'Completed_Orders/'
        self.original_orders_path = 'Original_Orders/'
        self.current_file = ""
        self.images = []
        self.page_count = 0
        
        self.paper = ""
        self.parent_sheet_size = ""
        self.parent_sheet_quantity = ""
        self.press_sheet_size = ""
        self.press_sheet_quantity = ""
        self.imposition = ""
        self.press = ""
        self.bindery = ""
        self.bindery_other = ""
        self.mail_department = ""
        self.display_graphics = ""
        self.po_number = ""
        
        self.configurations = json.loads(open('vault/configuration.json', 'r').read())
     
    def run_gui(self):
        self.count_files()

        dpg.create_context()
        dpg.create_viewport()
        dpg.setup_dearpygui()
        #dpg.toggle_viewport_fullscreen()
        
        if self.file_count == 0:
            self.no_orders_window()
        else:
            self.current_file = os.listdir(self.orders_path)[0]
            self.image_convert()
            
            self.create_window()

            dpg.show_viewport()
            dpg.start_dearpygui()
            dpg.destroy_context()
        
    def no_orders_window(self):
        with dpg.window(tag="No Orders"):
            dpg.add_text("There are no order forms in the Orders_to_Process folder. Please add at least one PDF and try again.", parent="No Orders")
            dpg.add_button(label="Exit", callback=self.exit_callback, parent="No Orders")
            dpg.set_primary_window("No Orders", True)
            dpg.show_viewport()
            dpg.start_dearpygui()
            dpg.destroy_context()
    def image_convert(self):
        self.images = convert_from_path((self.orders_path + self.current_file), fmt='png')
        x = 0
        for image in self.images:
            image.save(self.orders_path + str(x) + ".png")

            width, height, channels, data = dpg.load_image(self.orders_path + (str(x) + '.png'))
            with dpg.texture_registry():
                dpg.add_static_texture(width, height, data, tag=(str(self.current_file) + str(x) + ".png"))

            x+=1
            
        pdf_reader = PdfReader(str(self.orders_path + self.current_file))
        output_file_path = self.completed_orders_path + (str(self.current_file.split('.')[0]) + '.txt')
        page = pdf_reader.pages[0]
        try:
            result = page.extract_text().split("PURCHASE ORDER:")[1]
        except:
            result = "No PO Number Found"
        self.po_number = result.split()[0]
        self.page_count = len(pdf_reader.pages)
        
    def create_window(self):
        with dpg.window(tag="Order Form"):
            paper = dpg.add_group(horizontal=True)
            dpg.add_text("Paper:", parent=paper)
            dpg.add_combo(items=(self.configurations["Paper"]), callback=self.paper_callback, parent=paper, width=200)
            dpg.add_text("Other:", parent=paper)
            dpg.add_input_text(callback=self.paper_callback, parent=paper, width=200)
            
            parent_sheet_size = dpg.add_group(horizontal=True)
            dpg.add_text("Parent Sheet Size:", parent=parent_sheet_size)
            dpg.add_combo(items=(self.configurations["Parent Sheet Size"]), callback=self.parent_sheet_size_callback, parent=parent_sheet_size, width=200)
            dpg.add_text("Other:", parent=parent_sheet_size)
            dpg.add_input_text(callback=self.parent_sheet_size_callback, parent=parent_sheet_size, width=200)
            
            press_sheet_size = dpg.add_group(horizontal=True)
            dpg.add_text("Press Sheet Size:", parent=press_sheet_size)
            dpg.add_combo(items=(self.configurations["Press Sheet Size"]), callback=self.press_sheet_size_callback, parent=press_sheet_size, width=200)
            dpg.add_text("Other:", parent=press_sheet_size)
            dpg.add_input_text(callback=self.press_sheet_size_callback, parent=press_sheet_size, width=200)
            
            press_sheet_quantity = dpg.add_group(horizontal=True)
            dpg.add_text("Press Sheet Quantity:", parent=press_sheet_quantity)
            dpg.add_input_text(callback=self.press_sheet_quantity_callback, parent=press_sheet_quantity, width=200)

            imposition = dpg.add_group(horizontal=True)
            dpg.add_text("Imposition:", parent=imposition)
            dpg.add_combo(items=(self.configurations["Imposition"]), callback=self.imposition_callback, parent=imposition, width=200)
            dpg.add_text("Other:", parent=imposition)
            dpg.add_input_text(callback=self.imposition_callback, parent=imposition, width=200)
            
            press = dpg.add_group(horizontal=True)
            dpg.add_text("Press:", parent=press)
            dpg.add_combo(items=(self.configurations["Press"]), callback=self.press_callback, parent=press, width=200)
            dpg.add_text("Other:", parent=press)
            dpg.add_input_text(callback=self.press_callback, parent=press, width=200)
            
            mail_department = dpg.add_group(horizontal=True)
            dpg.add_text("Mail Department:", parent=mail_department)
            dpg.add_combo(items=(self.configurations["Mail Department"]), callback=self.mail_department_callback, parent=mail_department, width=200)
            dpg.add_text("Other:", parent=mail_department)
            dpg.add_input_text(callback=self.mail_department_callback, parent=mail_department, width=200)
            
            display_graphics = dpg.add_group(horizontal=True)
            dpg.add_text("Display Graphics:", parent=display_graphics)
            dpg.add_combo(items=(self.configurations["Display Graphics"]), callback=self.display_graphics_callback, parent=display_graphics, width=200)
            dpg.add_text("Other:", parent=display_graphics)
            dpg.add_input_text(callback=self.display_graphics_callback, parent=display_graphics, width=200)
            
            bindery = dpg.add_group(horizontal=True)
            dpg.add_text("Bindery:", parent=bindery)
            for item in self.configurations["Bindery"]:
                dpg.add_checkbox(label=item, callback=self.bindery_callback, user_data=item, parent=bindery)
            dpg.add_text("Other:", parent=bindery)
            dpg.add_input_text(callback=self.other_bindery_callback, parent=bindery, width=200)
            
            buttons = dpg.add_group(horizontal=True)
            dpg.add_button(label="Save", callback=self.save_confirmation_callback, parent=buttons)
            dpg.add_button(label="Exit", callback=self.exit_callback, parent=buttons)
            
            for x in range(len(self.images)):
                with dpg.drawlist(width=1000, height=1000, tag=("Order Window " + str(self.current_file) + str(x))):
                    dpg.draw_image((str(self.current_file) + str(x) + ".png"), (0, 0), (1000, 1000), uv_min=(0, 0), uv_max=(1, 1), tag=("Order Image " + str(self.current_file) + str(x)))
                x+=1
            
        dpg.set_primary_window("Order Form", True)
        
    def count_files(self):
        try:
            for path in os.listdir(self.orders_path):
                ext = os.path.splitext(path)[-1].lower()
                if os.path.isfile(os.path.join(self.orders_path, path)) and ext == ".pdf":
                    self.file_count += 1
                elif ext == ".png":
                    os.remove(self.orders_path + path)
        except:
            self.no_orders_window()
        
    def exit_callback(self):
        dpg.stop_dearpygui()
        
    def back_callback(self):
        dpg.delete_item("Order Inputs")
        dpg.delete_item("Order Form")
        dpg.delete_item("Back Button")
        self.create_window()
        self.clear_parameters()
        
    def save_confirmation_callback(self):
        self.bindery = self.bindery + self.bindery_other
        if self.bindery[-2] == ",":
            self.bindery = self.bindery[:-2]
        self.set_parent_sheet_quantity()
        string = "RUN PLAN - PURCHASE ORDER " + self.po_number + "\n\nPaper: " + self.paper + "\nParent Sheet Size: " + self.parent_sheet_size + "\nParent Sheet Quantity: " + self.parent_sheet_quantity + "\nPress Sheet Size: " + self.press_sheet_size + "\nPress Sheet Quantity: " + self.press_sheet_quantity + "\nImposition: " + self.imposition + "\nPress: " + self.press + "\nBindery: " + self.bindery + "\nMail Department: " + self.mail_department + "\nDisplay Graphics: " + self.display_graphics

        with dpg.window(tag="Back Button"):
            dpg.add_text("Press Continue to annotate PDF with the values below. Press Back to revise.", parent="Back Button")
            dpg.add_text(string, parent="Back Button")
            dpg.add_button(label="Continue", callback=self.save_callback, parent="Back Button")
            dpg.add_button(label="Back", callback=self.back_callback, parent="Back Button")
            
    def save_callback(self):
        string = "RUN PLAN - PURCHASE ORDER " + self.po_number + "\n\nPaper: " + self.paper + "\nParent Sheet Size: " + self.parent_sheet_size + "\nParent Sheet Quantity: " + self.parent_sheet_quantity + "\nPress Sheet Size: " + self.press_sheet_size + "\nPress Sheet Quantity: " + self.press_sheet_quantity + "\nImposition: " + self.imposition + "\nPress: " + self.press + "\nBindery: " + self.bindery + "\nMail Department: " + self.mail_department + "\nDisplay Graphics: " + self.display_graphics

        reader = PdfReader(str(self.orders_path + self.current_file))
        writer = PdfWriter()
        for x in range(0, self.page_count):
            writer.add_page(reader.pages[x])
        writer.add_blank_page(width=612, height=792)

        annotation = AnnotationBuilder.free_text(
            string,
            rect=(50, 396, 612, 792),
            font="Arial",
            bold=True,
            italic=True,
            font_size="20pt",
            font_color="000000",
            border_color="0000ff",
            background_color="ffffff",
        )
        writer.add_annotation(page_number=-1, annotation=annotation)
        
        with open((str(self.completed_orders_path + self.current_file.split('.')[0]) + "_annotated.pdf"), "wb") as fp:
            writer.write(fp)
        try:
            dpg.delete_item("Order Inputs")
            dpg.delete_item("Order Form")
            dpg.delete_item("Back Button")
            self.create_window()
        finally:
            for x in range(len(self.images)):
                dpg.delete_item("Order Image " + str(self.current_file) + str(x))
                dpg.delete_item("Order Window " + str(self.current_file) + str(x))
            for x in range(len(self.images)):
                os.remove(self.orders_path + (str(x) + '.png'))
            os.replace((self.orders_path + self.current_file), (self.original_orders_path + self.current_file))
            self.file_count -= 1
            if self.file_count > 0:
                self.current_file = os.listdir(self.orders_path)[0]
                self.image_convert()
                for x in range(len(self.images)):
                    with dpg.drawlist(width=1000, height=1000, parent="Order Form", tag=("Order Window " + str(self.current_file) + str(x))):
                        dpg.draw_image((str(self.current_file) + str(x) + ".png"), (0, 0), (1000, 1000), uv_min=(0, 0), uv_max=(1, 1), tag=("Order Image " + str(self.current_file) + str(x)))
                    x+=1
                self.clear_parameters()
            else:
                dpg.add_text("No orders left to process", parent="Order Form", pos=[400,500])
                
    def clear_parameters(self):
        self.paper = ""
        self.parent_sheet_size = ""
        self.parent_sheet_quantity = ""
        self.press_sheet_size = ""
        self.press_sheet_quantity = ""
        self.imposition = ""
        self.press = ""
        self.bindery = ""
        self.bindery_other = ""
        self.mail_department = ""
        self.display_graphics = ""
        
    def set_parent_sheet_quantity(self):
        try:
            parent_height = float(self.parent_sheet_size.split('x')[1])
            parent_width = float(self.parent_sheet_size.split('x')[0])
            press_height = float(self.press_sheet_size.split('x')[1])
            press_width = float(self.press_sheet_size.split('x')[0])
            height = parent_height // press_height
            width = parent_width // press_width
            if height < width:
                multiplier = height
            else:
                multiplier = width
            self.parent_sheet_quantity = str(int(self.press_sheet_quantity) * int(multiplier))
        except:
            self.parent_sheet_quantity = "Please enter sheet sizes and press sent quantity and save again."
                   
    def paper_callback(self, sender, data):
        self.paper = data
                    
    def parent_sheet_size_callback(self, sender, data):
        self.parent_sheet_size = data
                             
    def press_sheet_size_callback(self, sender, data):
        self.press_sheet_size = data
                         
    def press_sheet_quantity_callback(self, sender, data):
        self.press_sheet_quantity = data
                        
    def imposition_callback(self, sender, data):
        self.imposition = data
     
    def press_callback(self, sender, data):
        self.press = data
            
    def bindery_callback(self, sender, data, user_data):
        if data == True:
            self.bindery = self.bindery + user_data + ", "
        elif data == False:
            self.bindery.replace((user_data + ", "), "")
            
    def other_bindery_callback(self, sender, data, user_data):
        self.bindery_other = data
        
    def mail_department_callback(self, sender, data):
        self.mail_department = data
             
    def display_graphics_callback(self, sender, data):
        self.display_graphics = data
        
a = OrderForm()
a.run_gui()