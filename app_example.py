import dearpygui.dearpygui as dpg
from pdf2image import convert_from_path
import os
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import AnnotationBuilder

class OrderForm:

    def __init__(self):
        self.file_count = 0
        self.orders_path = 'Orders_to_Process/'
        self.completed_orders_path = 'Completed_Orders/'
        self.original_orders_path = 'Original_Orders/'
        self.current_file = ""
        self.machine = ""
        self.paper = ""
        self.parent_sheet_size = ""
        self.parent_sheet_quantity = ""
        self.press_sheet_size = ""
        self.press_sheet_quantity = ""
     

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

        def save_callback():
            string = self.current_file + " Saved \n" + self.machine + "\n" + self.paper + "\n" + self.parent_sheet_size + "\n" + self.parent_sheet_quantity + "\n" + self.press_sheet_size + "\n" + self.press_sheet_quantity
                    # Fill the writer with the pages you want
            print(str(self.orders_path + self.current_file))
            reader = PdfReader(str(self.orders_path + self.current_file))
            page = reader.pages[0]
            writer = PdfWriter()
            writer.add_page(page)

            # Create the annotation and add it
            annotation = AnnotationBuilder.free_text(
                string,
                rect=(50, 50, 200, 150),
                font="Arial",
                bold=True,
                italic=True,
                font_size="20pt",
                font_color="00ff00",
                border_color="0000ff",
                background_color="cdcdcd",
            )
            writer.add_annotation(page_number=0, annotation=annotation)

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
                    print(page.extract_text())
      
                    with dpg.texture_registry():
                        dpg.add_static_texture(width, height, data, tag=self.current_file)
                    with dpg.drawlist(width=500, height=500, parent="Order Form", tag="Order Window"):
                        dpg.draw_image(self.current_file, (0, 0), (500, 500), uv_min=(0, 0), uv_max=(1, 1), tag="Order Image", parent="Order Image")
                else:
                    dpg.add_text("No orders left to process", parent="Order Form", pos=[500,300])
                
        def exit_callback():
            dpg.stop_dearpygui()
            
        def machine_callback(sender, data):
            self.machine = data
            
        def paper_callback(sender, data):
            self.paper = data
            
        def parent_sheet_size_callback(sender, data):
            self.parent_sheet_size = data
            
        def parent_sheet_quantity_callback(sender, data):
            self.parent_sheet_quantity = data
            
        def press_sheet_size_callback(sender, data):
            self.press_sheet_size = data
            
        def press_sheet_quantity_callback(sender, data):
            self.press_sheet_quantity = data
                 
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
        print(page.extract_text())
        
        with dpg.texture_registry():
            dpg.add_static_texture(width, height, data, tag=self.current_file)
        
        with dpg.window(tag="Order Form"):
            dpg.add_combo(label="Machine", items=("machine1", "machine2", "machine3"), callback=machine_callback)
            dpg.add_combo(label="Paper", items=("paper1", "paper2", "paper3"), callback=paper_callback)
            dpg.add_combo(label="Parent Sheet Size", items=("size1", "size2", "size3"), callback=parent_sheet_size_callback)
            dpg.add_combo(label="Parent Sheet Quantity", items=("1", "2", "3"), callback=parent_sheet_quantity_callback)
            dpg.add_combo(label="Press Sheet Size", items=("size1", "size2", "size3"), callback=press_sheet_size_callback)
            dpg.add_combo(label="Press Sheet Quantity", items=("1", "2", "3"), callback=press_sheet_quantity_callback)
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
