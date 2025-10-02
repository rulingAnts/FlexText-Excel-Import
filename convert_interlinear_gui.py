import tkinter as tk
from tkinter import ttk, filedialog
import os
from xml.etree.ElementTree import tostring
from xml.dom import minidom

from excel_to_xml import convert_excel_to_xml_dom
from xml_to_flextext import transform_to_flextext_dom

# TO DO: make the specific conversion engines into classes 
# and structure operations so they don't block Tkinter's event loop for long.
# tkdocs.com/tutorial/eventloop.html

class Converter(tk.Tk):
    def __init__(self):
        """
        GUI application for converting interlinear data.

        __init__ method sets up the main window and all widgets.
        """
        super().__init__()
        self.title("Interlinear Converter")
        self.intermediate_xml = None
        self.data_loaded = False

        self.mainframe = ttk.Frame(self, padding="10 10 10 10")
        self.mainframe.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))

        # Input block
        self.inputFormatLabel = ttk.Label(self.mainframe, text="Input Format:")
        self.inputFormatLabel.grid(row=0, column=0, pady=5, padx=5)
        self.inputFormatCombo = ttk.Combobox(self.mainframe, values=["Excel Interlinear"])
        self.inputFormatCombo.bind('<<ComboboxSelected>>', lambda e: self.inputLoadButton.state(['!disabled']))
        self.inputFormatCombo.grid(row=0, column=1, pady=5, padx=5)
        self.inputLoadButton = ttk.Button(self.mainframe, text="Select input file & load", state='disabled', command=self.load_file)
        self.inputLoadButton.grid(row=0, column=2, pady=5, padx=5)
        self.loadProgress = ttk.Progressbar(self.mainframe, orient=tk.HORIZONTAL, mode='indeterminate')
        self.loadProgress.grid(row=1, column=0, columnspan=3, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.hide_load_progress()

        # Reminders & writing systems
        self.reminderLabel = ttk.Label(self.mainframe, text="Reminders about setting language and writing systems in Excel and in FLEx.")
        self.reminderLabel.grid(row=3, column=0, columnspan=3, pady=5, padx=5)
        self.reminderLabel.config(wraplength=400)
        self.wsVernacularLabel = ttk.Label(self.mainframe, text="Vernacular writing system:", anchor='e')
        self.wsVernacularLabel.grid(row=4, column=0, pady=1, padx=5)
        self.wsVernacular = ttk.Label(self.mainframe, text="(not loaded)", anchor='w')
        self.wsVernacular.grid(row=4, column=1, pady=1, padx=5)
        self.wsFreeLabel = ttk.Label(self.mainframe, text="Free trans. writing system:", anchor='e')
        self.wsFreeLabel.grid(row=5, column=0, pady=1, padx=5)
        self.wsFree = ttk.Label(self.mainframe, text="(not loaded)", anchor='w')
        self.wsFree.grid(row=5, column=1, pady=1, padx=5)
        self.wsGlossLabel = ttk.Label(self.mainframe, text="Gloss writing system:", anchor='e')
        self.wsGlossLabel.grid(row=6, column=0, pady=1, padx=5)
        self.wsGloss = ttk.Label(self.mainframe, text="(not loaded)", anchor='w')
        self.wsGloss.grid(row=6, column=1, pady=1, padx=5)

        # Output block
        self.outputFormatLabel = ttk.Label(self.mainframe, text="Output Format:")
        self.outputFormatLabel.grid(row=8, column=0, pady=5, padx=5)
        self.outputFormatCombo = ttk.Combobox(self.mainframe, values=["FlexText Interlinear"])
        self.outputFormatCombo.bind('<<ComboboxSelected>>', self.update_convert_button_state)
        self.outputFormatCombo.grid(row=8, column=1, pady=5, padx=5)
        self.convertButton = ttk.Button(self.mainframe, text="Select output file & convert", state='disabled', command=self.convert)
        self.convertButton.grid(row=8, column=2, pady=5, padx=5)
        self.convertProgress = ttk.Progressbar(self.mainframe, orient=tk.HORIZONTAL, mode='indeterminate')
        self.convertProgress.grid(row=9, column=0, columnspan=3, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.hide_convert_progress()

        # Error display
        self.errorDisplay = ttk.Label(self.mainframe, text="", foreground="red", wraplength=400)
        self.errorDisplay.grid(row=10, column=0, columnspan=3, pady=5, padx=5)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    def update_convert_button_state(self):
        """
        Enable the convert button if input is loaded and output format is selected.
        """
        
        if self.data_loaded and self.outputFormatCombo.get():
            self.convertButton.state(['!disabled'])
        else:
            self.convertButton.state(['disabled'])

    def update_writing_systems(self):
        if self.data_loaded:
            metadata = self.intermediate_xml.find('text_metadata')
            if metadata is not None:
                metadataVernacular = metadata.find('writing_system_vernacular')
                metadataGloss = metadata.find('writing_system_gloss')
                metadataFree = metadata.find('writing_system_free')
                if metadataVernacular is not None:
                    self.wsVernacular.config(text=metadataVernacular.text)
                if metadataGloss is not None:
                    self.wsGloss.config(text=metadataGloss.text)
                if metadataFree is not None:
                    self.wsFree.config(text=metadataFree.text)
            else:
                # log error
                # TODO: how to do this cleanly? function clear_writing_systems() or better if/else logic?
                self.wsVernacular.config(text="(not loaded)")
                self.wsGloss.config(text="(not loaded)")
                self.wsFree.config(text="(not loaded)")
        else:
            self.wsVernacular.config(text="(not loaded)")
            self.wsGloss.config(text="(not loaded)")
            self.wsFree.config(text="(not loaded)")

    def hide_load_progress(self):
        self.loadProgress.grid_remove()

    def show_load_progress(self):
        self.loadProgress.grid()

    def hide_convert_progress(self):
        self.convertProgress.grid_remove()

    def show_convert_progress(self):
        self.convertProgress.grid()

    def load_file(self):
        """
        Loads a file from a file dialog and processes it into intermediate XML.
        """

        formatString = self.inputFormatCombo.get()
        if formatString == "Excel Interlinear":
            filetypelist = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        elif formatString == "XLingPaper Interlinear":
            filetypelist = [("XLingPaper files", "*.xml"), ("All files", "*.*")]
        else:
            raise ValueError("Unsupported input format selected. Add code here")
        
        filepath = filedialog.askopenfilename(title="Select a file", filetypes=filetypelist)
        if not filepath:
            return None     # User cancelled
        if not os.path.exists(filepath):
            self.errorDisplay.config(text="Error: File does not exist.")
            return None

        self.show_load_progress()
        self.loadProgress.start()
        if formatString == "Excel Interlinear":
            try:
                (self.intermediate_xml, error_list) = convert_excel_to_xml_dom(filepath)
            except Exception as e:
                # most exceptions are reported via error_list, currently.
                # TODO
                pass
            # TODO: set writing system labels
            # self.inputLang.config(text="Language: [set after loading file]")
        # check that writing systems are defined
        
        self.loadProgress.stop()
        self.hide_load_progress()
        self.data_loaded = True
        self.update_convert_button_state()

    def convert(self):
        """
        Sets an output file from a file dialog and converts into target format.
        """

        formatString = self.outputFormatCombo.get()
        if formatString == "FlexText Interlinear":
            filetypelist = [("FlexText files", "*.flextext"), ("All files", "*.*")]
        elif formatString == "XLingPaper Interlinear":
            filetypelist = [("XLingPaper files", "*.xml"), ("All files", "*.*")]
        else:
            raise ValueError("Unsupported output format selected. Add code here")
        filenamebase, _ = os.path.splitext(self.inputFileName.cget("text"))
        initialfilename = filenamebase + ".flextext"
        filepath = filedialog.asksaveasfilename(initialfile=initialfilename, defaultextension=".flextext")
        if filepath:
            _, outputExt = os.path.splitext(filepath)
            if outputExt.lower() in [".flextext"]:
                self.outputFormatCombo.set("FlexText Interlinear")
            self.outputFileName.config(text=filepath)
            self.convertButton.state(['!disabled'])

        self.show_convert_progress()
        self.convertProgress.start()
        try:
            (flextext_xml, missing_freetrans_count) = transform_to_flextext_dom(
                self.intermediate_xml, self.wsVernacular.get(), self.wsGloss.get(), self.wsFree.get())
        except Exception:
            # check what exceptions this function might raise
            # TODO
            pass
        pretty_xml = self.prettify_xml(flextext_xml)
        with open(self.outputFileName.cget("text"), 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
        self.convertProgress.stop()
        self.hide_convert_progress()

    def prettify_xml(self, element):
        """
        Add whitespace to make an XML element tree pretty-printed.
        """

        rough_string = tostring(element, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(encoding='utf-8', indent="  ").decode('utf-8')


if __name__ == "__main__":
    app = Converter()
    app.mainloop()