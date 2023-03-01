import dearpygui.dearpygui as dpg
import json
import os
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import AnnotationBuilder
from pdf2image import convert_from_path
import requests
import shutil
import sys
import webbrowser
import zipfile

__name__ = 'PDF Annotation App'
__version__ = '1.1'
__author__ = 'Darby Hudnall'
__email__ = 'DarbyHudnall95@gmail.com'

class OrderForm:

    def __init__(self):
        self.file_count = 0
        self.orders_path = 'Orders_to_Process/'
        self.completed_orders_path = 'Completed_Orders/'
        self.original_orders_path = 'Original_Orders/'
        self.current_file = ""
        self.images = []
        self.page_count = 0
        
        self.cover_paper = ""
        self.cover_parent_sheet_size = ""
        self.cover_parent_sheet_quantity = ""
        self.cover_press_sheet_size = ""
        self.cover_press_sheet_quantity = ""
        self.text_paper = ""
        self.text_parent_sheet_size = ""
        self.text_parent_sheet_quantity = ""
        self.text_press_sheet_size = ""
        self.text_press_sheet_quantity = ""
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
        
        self.check_updates()
        
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
            
    def upgrade_needed_window(self, message):
        with dpg.window(tag="Upgrade Needed"):
            dpg.add_text(message, parent="Upgrade Needed")
            dpg.add_text(os.path.dirname(os.path.abspath(sys.argv[0])))
            dpg.add_button(label="Complete Upgrade", callback=self.upgrade_callback, parent="Upgrade Needed")
            dpg.add_button(label="Exit", callback=self.exit_callback, parent="Upgrade Needed")
            dpg.set_primary_window("Upgrade Needed", True)
            dpg.show_viewport()
            dpg.start_dearpygui()
            dpg.destroy_context()
           
    def image_convert(self):
        #self.images = convert_from_path((self.orders_path + self.current_file), fmt='png') #mac version
        self.images = convert_from_path((self.orders_path + self.current_file), fmt='png', poppler_path=r"C:\Program Files\poppler-23.01.0\Library\bin") #windows version
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
        
    def check_updates(self):
        response = requests.get(
            'https://raw.githubusercontent.com/dhudnall26/desktop_app_project/master/version.txt')
        data = response.text

        if float(data) > float(__version__):
            message = "Need to upgrade " + __name__ + " Version " + __version__ + " to Version " + data
            self.upgrade_needed_window(message)

    def create_window(self):
        with dpg.window(tag="Order Form"):
            cover_paper = dpg.add_group(horizontal=True)
            dpg.add_text("Paper Cover:", parent=cover_paper)
            dpg.add_combo(items=(self.configurations["Paper Cover"]), callback=self.cover_paper_callback, parent=cover_paper, width=200)
            dpg.add_text("Other:", parent=cover_paper)
            dpg.add_input_text(callback=self.cover_paper_callback, parent=cover_paper, width=200)
            
            cover_parent_sheet_size = dpg.add_group(horizontal=True)
            dpg.add_text("Cover Parent Sheet Size:", parent=cover_parent_sheet_size)
            dpg.add_combo(items=(self.configurations["Cover Parent Sheet Size"]), callback=self.cover_parent_sheet_size_callback, parent=cover_parent_sheet_size, width=200)
            dpg.add_text("Other:", parent=cover_parent_sheet_size)
            dpg.add_input_text(callback=self.cover_parent_sheet_size_callback, parent=cover_parent_sheet_size, width=200)
            
            cover_parent_sheet_quantity = dpg.add_group(horizontal=True)
            dpg.add_text("Cover Parent Sheet Quantity:", parent=cover_parent_sheet_quantity)
            dpg.add_input_text(callback=self.cover_parent_sheet_quantity_callback, parent=cover_parent_sheet_quantity, width=200)

            cover_press_sheet_size = dpg.add_group(horizontal=True)
            dpg.add_text("Cover Press Sheet Size:", parent=cover_press_sheet_size)
            dpg.add_combo(items=(self.configurations["Cover Press Sheet Size"]), callback=self.cover_press_sheet_size_callback, parent=cover_press_sheet_size, width=200)
            dpg.add_text("Other:", parent=cover_press_sheet_size)
            dpg.add_input_text(callback=self.cover_press_sheet_size_callback, parent=cover_press_sheet_size, width=200)
            
            cover_press_sheet_quantity = dpg.add_group(horizontal=True)
            dpg.add_text("Cover Press Sheet Quantity:", parent=cover_press_sheet_quantity)
            dpg.add_input_text(callback=self.cover_press_sheet_quantity_callback, parent=cover_press_sheet_quantity, width=200)

            text_paper = dpg.add_group(horizontal=True)
            dpg.add_text("Paper Text:", parent=text_paper)
            dpg.add_combo(items=(self.configurations["Paper Text"]), callback=self.text_paper_callback, parent=text_paper, width=200)
            dpg.add_text("Other:", parent=text_paper)
            dpg.add_input_text(callback=self.text_paper_callback, parent=text_paper, width=200)
            
            text_parent_sheet_size = dpg.add_group(horizontal=True)
            dpg.add_text("Text Parent Sheet Size:", parent=text_parent_sheet_size)
            dpg.add_combo(items=(self.configurations["Text Parent Sheet Size"]), callback=self.text_parent_sheet_size_callback, parent=text_parent_sheet_size, width=200)
            dpg.add_text("Other:", parent=text_parent_sheet_size)
            dpg.add_input_text(callback=self.text_parent_sheet_size_callback, parent=text_parent_sheet_size, width=200)
            
            text_parent_sheet_quantity = dpg.add_group(horizontal=True)
            dpg.add_text("Text Parent Sheet Quantity:", parent=text_parent_sheet_quantity)
            dpg.add_input_text(callback=self.text_parent_sheet_quantity_callback, parent=text_parent_sheet_quantity, width=200)

            text_press_sheet_size = dpg.add_group(horizontal=True)
            dpg.add_text("Text Press Sheet Size:", parent=text_press_sheet_size)
            dpg.add_combo(items=(self.configurations["Text Press Sheet Size"]), callback=self.text_press_sheet_size_callback, parent=text_press_sheet_size, width=200)
            dpg.add_text("Other:", parent=text_press_sheet_size)
            dpg.add_input_text(callback=self.text_press_sheet_size_callback, parent=text_press_sheet_size, width=200)
            
            text_press_sheet_quantity = dpg.add_group(horizontal=True)
            dpg.add_text("Text Press Sheet Quantity:", parent=text_press_sheet_quantity)
            dpg.add_input_text(callback=self.text_press_sheet_quantity_callback, parent=text_press_sheet_quantity, width=200)

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
        
    def upgrade_callback(self):
        #webbrowser.open_new_tab('https://github.com/dhudnall26/desktop_app_project/tree/master')
        r = requests.get('https://raw.githubusercontent.com/dhudnall26/desktop_app_project/master/pdf_annotation_app.zip', stream=True)
        with open('pdf_annotation_app.zip', 'wb') as f:
            shutil.copyfileobj(r.raw, f)
            
        with zipfile.ZipFile('pdf_annotation_app.zip', 'r') as zip_ref:
            zip_ref.extractall(os.path.dirname(os.path.abspath(sys.argv[0])))
            #zip_ref.extractall('/Users/DarbyHudnall/order_form')
            
        with dpg.window(tag="Upgrade Complete"):
            dpg.add_text("Update complete. Press exit then open the app again to use the latest version.", parent="Upgrade Complete")
            dpg.add_button(label="Exit", callback=self.exit_callback, parent="Upgrade Complete")
        
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
        #self.set_parent_sheet_quantity()
        string = "RUN PLAN - PURCHASE ORDER " + self.po_number + "\n\nPaper Cover: " + self.cover_paper + "\nCover Parent Sheet Size: " + self.cover_parent_sheet_size + "\nCover Parent Sheet Quantity: " + self.cover_parent_sheet_quantity + "\nCover Press Sheet Size: " + self.cover_press_sheet_size + "\nCover Press Sheet Quantity: " + self.cover_press_sheet_quantity + "\nPaper Text: " + self.text_paper + "\nText Parent Sheet Size: " + self.text_parent_sheet_size + "\nText Parent Sheet Quantity: " + self.text_parent_sheet_quantity + "\nText Press Sheet Size: " + self.text_press_sheet_size + "\nText Press Sheet Quantity: " + self.text_press_sheet_quantity + "\nImposition: " + self.imposition + "\nPress: " + self.press + "\nBindery: " + self.bindery + "\nMail Department: " + self.mail_department + "\nDisplay Graphics: " + self.display_graphics

        with dpg.window(tag="Back Button"):
            dpg.add_text("Press Continue to annotate PDF with the values below. Press Back to revise.", parent="Back Button")
            dpg.add_text(string, parent="Back Button")
            dpg.add_button(label="Continue", callback=self.save_callback, parent="Back Button")
            dpg.add_button(label="Back", callback=self.back_callback, parent="Back Button")
            
    def save_callback(self):
        string = "RUN PLAN - PURCHASE ORDER " + self.po_number + "\n\nPaper Cover: " + self.cover_paper + "\nCover Parent Sheet Size: " + self.cover_parent_sheet_size + "\nCover Parent Sheet Quantity: " + self.cover_parent_sheet_quantity + "\nCover Press Sheet Size: " + self.cover_press_sheet_size + "\nCover Press Sheet Quantity: " + self.cover_press_sheet_quantity + "\nPaper Text: " + self.text_paper + "\nText Parent Sheet Size: " + self.text_parent_sheet_size + "\nText Parent Sheet Quantity: " + self.text_parent_sheet_quantity + "\nText Press Sheet Size: " + self.text_press_sheet_size + "\nText Press Sheet Quantity: " + self.text_press_sheet_quantity + "\nImposition: " + self.imposition + "\nPress: " + self.press + "\nBindery: " + self.bindery + "\nMail Department: " + self.mail_department + "\nDisplay Graphics: " + self.display_graphics

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
        self.cover_paper = ""
        self.cover_parent_sheet_size = ""
        self.cover_parent_sheet_quantity = ""
        self.cover_press_sheet_size = ""
        self.cover_press_sheet_quantity = ""
        self.text_paper = ""
        self.text_parent_sheet_size = ""
        self.text_parent_sheet_quantity = ""
        self.text_press_sheet_size = ""
        self.text_press_sheet_quantity = ""
        self.imposition = ""
        self.press = ""
        self.bindery = ""
        self.bindery_other = ""
        self.mail_department = ""
        self.display_graphics = ""
        
    # def set_parent_sheet_quantity(self):
    #     try:
    #         parent_height = float(self.parent_sheet_size.split('x')[1])
    #         parent_width = float(self.parent_sheet_size.split('x')[0])
    #         press_height = float(self.press_sheet_size.split('x')[1])
    #         press_width = float(self.press_sheet_size.split('x')[0])
    #         height = parent_height // press_height
    #         width = parent_width // press_width
    #         if height < width:
    #             multiplier = height
    #         else:
    #             multiplier = width
    #         self.parent_sheet_quantity = str(int(self.press_sheet_quantity) * int(multiplier))
    #     except:
    #         self.parent_sheet_quantity = "Please enter sheet sizes and press sent quantity and save again."
                   
    def cover_paper_callback(self, sender, data):
        self.cover_paper = data
                    
    def cover_parent_sheet_size_callback(self, sender, data):
        self.cover_parent_sheet_size = data
    
    def cover_parent_sheet_quantity_callback(self, sender, data):
        self.cover_press_sheet_quantity = data
                          
    def cover_press_sheet_size_callback(self, sender, data):
        self.cover_press_sheet_size = data
                         
    def cover_press_sheet_quantity_callback(self, sender, data):
        self.cover_press_sheet_quantity = data
    
    def text_paper_callback(self, sender, data):
        self.text_paper = data
                    
    def text_parent_sheet_size_callback(self, sender, data):
        self.text_parent_sheet_size = data
    
    def text_parent_sheet_quantity_callback(self, sender, data):
        self.text_press_sheet_quantity = data
                          
    def text_press_sheet_size_callback(self, sender, data):
        self.text_press_sheet_size = data
                         
    def text_press_sheet_quantity_callback(self, sender, data):
        self.text_press_sheet_quantity = data
                      
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
