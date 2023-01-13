import dearpygui.dearpygui as dpg
from pdf2image import convert_from_path
import os

class OrderForm:

    def __init__(self):
        self.file_count = 0
        self.orders_path = 'Orders/'
        self.completed_orders_path = 'Completed_Orders/'
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
            print('File count:', self.file_count)
        except:
            print("No orders to process.")
        
    def image_convert(self):
        images = convert_from_path(self.orders_path + self.current_file)
        for i in range(len(images)):
            images[i].save(self.orders_path + (str(self.current_file.split('.')[0]) + '.png'), 'PNG')
        
    def run_gui(self):

        def save_callback():
            string = "Order Saved \n" + self.machine + "\n" + self.paper + "\n" + self.parent_sheet_size + "\n" + self.parent_sheet_quantity + "\n" + self.press_sheet_size + "\n" + self.press_sheet_quantity
            try:
            	dpg.delete_item("Order Inputs")
            finally:
                #with dpg.drawlist(width=500, height=500):
                #    dpg.draw_image(self.current_file, (0, 0), (500, 500), uv_min=(0, 0), uv_max=(1, 1))
                dpg.add_text(string, parent="Order Form", pos=[600,0], tag="Order Inputs")
                os.remove(self.orders_path + (str(self.current_file.split('.')[0]) + '.png'))
                os.replace((self.orders_path + self.current_file), (self.completed_orders_path + self.current_file))
                self.file_count -= 1
                if self.file_count > 0:
                    self.current_file = os.listdir(self.orders_path)[0]
                else:
                    print("No orders left to process")
                    exit_callback()
        
                
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
            exit_callback()
        
        self.image_convert()
        width, height, channels, data = dpg.load_image(self.orders_path + (str(self.current_file.split('.')[0]) + '.png')) # 0: width, 1: height, 2: channels, 3: data
        
        with dpg.texture_registry():
            dpg.add_static_texture(width, height, data, tag=self.current_file)

        
        with dpg.window(tag="Order Form"):
            with dpg.drawlist(width=500, height=500):
                dpg.draw_image(self.current_file, (0, 0), (500, 500), uv_min=(0, 0), uv_max=(1, 1))
            dpg.add_combo(label="Machine", items=("machine1", "machine2", "machine3"), callback=machine_callback)
            dpg.add_combo(label="Paper", items=("paper1", "paper2", "paper3"), callback=paper_callback)
            dpg.add_combo(label="Parent Sheet Size", items=("size1", "size2", "size3"), callback=parent_sheet_size_callback)
            dpg.add_combo(label="Parent Sheet Quantity", items=("1", "2", "3"), callback=parent_sheet_quantity_callback)
            dpg.add_combo(label="Press Sheet Size", items=("size1", "size2", "size3"), callback=press_sheet_size_callback)
            dpg.add_combo(label="Press Sheet Quantity", items=("1", "2", "3"), callback=press_sheet_quantity_callback)
            dpg.add_button(label="Save", callback=save_callback)
            dpg.add_button(label="Exit", callback=exit_callback)
            
        dpg.set_primary_window("Order Form", True)
        dpg.show_viewport()
        dpg.start_dearpygui()
        dpg.destroy_context()
        
a = OrderForm()
a.run_gui()
