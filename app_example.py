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
        self.page_count = 0
        
        self.machine = ""
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
        
        self.configurations = json.loads(open('configuration.json', 'r').read())
     

    def count_files(self):
        try:
            for path in os.listdir(self.orders_path):
                ext = os.path.splitext(path)[-1].lower()
                if os.path.isfile(os.path.join(self.orders_path, path)) and ext == ".pdf":
                    self.file_count += 1
                elif ext == ".png":
                    os.remove(self.orders_path + path)
        except:
            print("No orders to process.")
        
    def image_convert(self):
        images = convert_from_path(self.orders_path + self.current_file)
        for i in range(len(images)):
            images[i].save(self.orders_path + (str(self.current_file.split('.')[0]) + '.png'), 'PNG')
        
    def run_gui(self):
        self.machine = ""
        self.paper = ""
        self.parent_sheet_size = ""
        self.parent_sheet_quantity = ""
        self.press_sheet_size = ""
        self.press_sheet_quantity = ""
        self.imposition = ""
        self.press = ""
        self.bindery = ""
        self.mail_department = ""
        self.display_graphics = ""
        
        def save_callback():
            self.bindery = self.bindery + self.bindery_other
            string = "RUN PLAN - PURCHASE ORDER " + self.po_number + "\n\nMachine: " + self.machine + "\nPaper: " + self.paper + "\nParent Sheet Size: " + self.parent_sheet_size + "\nParent Sheet Quantity: " + self.parent_sheet_quantity + "\nPress Sheet Size: " + self.press_sheet_size + "\nPress Sheet Quantity: " + self.press_sheet_quantity + "\nImposition: " + self.imposition + "\nPress: " + self.press + "\nBindery: " + self.bindery + "\nMail Department: " + self.mail_department + "\nDisplay Graphics: " + self.display_graphics
                    # Fill the writer with the pages you want
            print(str(self.orders_path + self.current_file))
            reader = PdfReader(str(self.orders_path + self.current_file))
            writer = PdfWriter()
            for x in range(0, self.page_count):
                writer.add_page(reader.pages[x])
            writer.add_blank_page(width=612, height=792)

            # Create the annotation and add it
            annotation = AnnotationBuilder.free_text(
                string,
                rect=(0, 396, 306, 792),
                font="Arial",
                bold=True,
                italic=True,
                font_size="20pt",
                font_color="000000",
                border_color="0000ff",
                background_color="ffffff",
            )
            writer.add_annotation(page_number=-1, annotation=annotation)

            # Write the annotated file to disk
            with open((str(self.completed_orders_path + self.current_file.split('.')[0]) + "_annotated.pdf"), "wb") as fp:
                writer.write(fp)
            try:
            	dpg.delete_item("Order Inputs")
            finally:
                dpg.delete_item("Order Image")
                dpg.delete_item("Order Window")
                dpg.add_text(string, parent="Order Form", pos=[600,400], tag="Order Inputs")
                os.remove(self.orders_path + (str(self.current_file.split('.')[0]) + '.png'))
                os.replace((self.orders_path + self.current_file), (self.original_orders_path + self.current_file))
                self.file_count -= 1
                if self.file_count > 0:
                    self.current_file = os.listdir(self.orders_path)[0]
                    self.image_convert()
                    width, height, channels, data = dpg.load_image(self.orders_path + (str(self.current_file.split('.')[0]) + '.png')) # 0: width, 1: height, 2: channels, 3: data
                    pdf_reader = PdfReader(str(self.orders_path + self.current_file))
                    output_file_path = self.completed_orders_path + (str(self.current_file.split('.')[0]) + '.txt')
                    
                    page = pdf_reader.pages[0]
                    result = page.extract_text().split("PURCHASE ORDER:")[1]
                    self.po_number = result.split()[0]
                    
                    self.page_count = len(pdf_reader.pages)
      
                    with dpg.texture_registry():
                        dpg.add_static_texture(width, height, data, tag=self.current_file)
                    with dpg.drawlist(width=500, height=500, parent="Order Form", tag="Order Window"):
                        dpg.draw_image(self.current_file, (0, 0), (500, 500), uv_min=(0, 0), uv_max=(1, 1), tag="Order Image", parent="Order Image")
                else:
                    dpg.add_text("No orders left to process", parent="Order Form", pos=[500,300])
                
        def exit_callback():
            dpg.stop_dearpygui()
            
        def machine_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Machine"):
                    dpg.add_input_text(label="Machine", width=500, height=500, callback=other_machine_callback)
            else:
                self.machine = data
                        
        def other_machine_callback(sender, data):
            self.machine = data
            
        def paper_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Paper"):
                    dpg.add_input_text(label="Paper", width=500, height=500, callback=other_paper_callback)
            else:
                self.paper = data
                        
        def other_paper_callback(sender, data):
            self.paper = data
            
        def parent_sheet_size_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Parent Sheet Size"):
                    dpg.add_input_text(label="Parent Sheet Size", width=500, height=500, callback=other_parent_sheet_size_callback)
            else:
                self.parent_sheet_size = data
                        
        def other_parent_sheet_size_callback(sender, data):
            self.parent_sheet_size = data
              
        def parent_sheet_quantity_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Parent Sheet Quantity"):
                    dpg.add_input_text(label="Parent Sheet Quantity", width=500, height=500, callback=other_parent_sheet_quantity_callback)
            else:
                self.parent_sheet_quantity = data
                        
        def other_parent_sheet_quantity_callback(sender, data):
            self.parent_sheet_quantity = data
              
        def press_sheet_size_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Press Sheet Size"):
                    dpg.add_input_text(label="Press Sheet Size", width=500, height=500, callback=other_press_sheet_size_callback)
            else:
                self.press_sheet_size = data
                        
        def other_press_sheet_size_callback(sender, data):
            self.press_sheet_size = data
             
        def press_sheet_quantity_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Press Sheet Quantity"):
                    dpg.add_input_text(label="Press Sheet Quantity", width=500, height=500, callback=other_press_sheet_quantity_callback)
            else:
                self.press_sheet_quantity = data
                        
        def other_press_sheet_quantity_callback(sender, data):
            self.press_sheet_quantity = data
             
        def imposition_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Imposition"):
                    dpg.add_input_text(label="Imposition", width=500, height=500, callback=other_imposition_callback)
            else:
                self.imposition = data
        
        def other_imposition_callback(sender, data):
            self.imposition = data
            
        def press_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Press"):
                    dpg.add_input_text(label="Press", width=500, height=500, callback=other_press_callback)
            else:
                self.press = data
                
        def other_press_callback(sender, data):
            self.press = data
            
        def bindery_callback(sender, data, user_data):
            if data == True:
                if user_data == "Other":
                    with dpg.window(tag="Other Bindery"):
                        dpg.add_input_text(label="Bindery", width=500, height=500, callback=other_bindery_callback, user_data=user_data)
                else:
                    self.bindery = self.bindery + user_data + ", "
            elif data == False:
                self.bindery.replace((user_data + ", "), "")
                
        def other_bindery_callback(sender, data, user_data):
            self.bindery_other = data
            
        def mail_department_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Mail Department"):
                    dpg.add_input_text(label="Mail Department", width=500, height=500, callback=other_mail_department_callback)
            else:
                self.mail_department = data
                
        def other_mail_department_callback(sender, data):
            self.mail_department = data
            
        def display_graphics_callback(sender, data):
            if data == "Other":
                with dpg.window(tag="Other Display Graphics"):
                    dpg.add_input_text(label="Display Graphics", width=500, height=500, callback=other_display_graphics_callback)
            else:
                self.display_graphics = data
                
        def other_display_graphics_callback(sender, data):
            self.display_graphics = data
                  
        self.count_files()

        dpg.create_context()
        dpg.create_viewport()
        dpg.setup_dearpygui()
        
        if self.file_count > 0:
            self.current_file = os.listdir(self.orders_path)[0]
        else:
            print("No orders left to process")
            exit()
        
        self.image_convert()
        width, height, channels, data = dpg.load_image(self.orders_path + (str(self.current_file.split('.')[0]) + '.png')) # 0: width, 1: height, 2: channels, 3: data
        
        pdf_reader = PdfReader(str(self.orders_path + self.current_file))
        output_file_path = self.completed_orders_path + (str(self.current_file.split('.')[0]) + '.txt')
        
        page = pdf_reader.pages[0]
        result = page.extract_text().split("PURCHASE ORDER:")[1]
        self.po_number = result.split()[0]
        
        self.page_count = len(pdf_reader.pages)

        with dpg.texture_registry():
            dpg.add_static_texture(width, height, data, tag=self.current_file)
        
        with dpg.window(tag="Order Form"):
            dpg.add_combo(label="Machine", items=(self.configurations["Machine"]), callback=machine_callback)
            dpg.add_combo(label="Paper", items=(self.configurations["Paper"]), callback=paper_callback)
            dpg.add_combo(label="Parent Sheet Size", items=(self.configurations["Parent Sheet Size"]), callback=parent_sheet_size_callback)
            dpg.add_combo(label="Parent Sheet Quantity", items=(self.configurations["Parent Sheet Quantity"]), callback=parent_sheet_quantity_callback)
            dpg.add_combo(label="Press Sheet Size", items=(self.configurations["Press Sheet Size"]), callback=press_sheet_size_callback)
            dpg.add_combo(label="Press Sheet Quantity", items=(self.configurations["Press Sheet Quantity"]), callback=press_sheet_quantity_callback)
            dpg.add_combo(label="Imposition", items=(self.configurations["Imposition"]), callback=imposition_callback)
            dpg.add_combo(label="Press", items=(self.configurations["Press"]), callback=press_callback)
            #dpg.add_combo(label="Bindery", items=(self.configurations["Bindery"]), callback=bindery_callback)
            dpg.add_text("Bindery")
            for item in self.configurations["Bindery"]:
                dpg.add_checkbox(label=item, callback=bindery_callback, user_data=item)
            dpg.add_combo(label="Mail Department", items=(self.configurations["Mail Department"]), callback=mail_department_callback)
            dpg.add_combo(label="Display Graphics", items=(self.configurations["Display Graphics"]), callback=display_graphics_callback)
            dpg.add_button(label="Save", callback=save_callback)
            dpg.add_button(label="Exit", callback=exit_callback)
            with dpg.drawlist(width=500, height=500, tag="Order Window"):
                dpg.draw_image(self.current_file, (0, 0), (500, 500), uv_min=(0, 0), uv_max=(1, 1), tag="Order Image")
            
        dpg.set_primary_window("Order Form", True)
        dpg.show_viewport()
        dpg.start_dearpygui()
        dpg.destroy_context()
        
a = OrderForm()
a.run_gui()
