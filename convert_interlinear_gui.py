import tkinter as tk
from tkinter import ttk, filedialog
import os
import traceback
from xml.etree.ElementTree import tostring
from xml.dom import minidom

from InterlinearLoaders import ExcelInterlinearLoader
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
        self.AFTER_DELAY_MS = 1 # milliseconds
        self.title("Interlinear Converter")
        self.intermediate_xml = None
        self.is_data_loaded = False
        self.writing_systems_ready = False
        self.inputFileName = None
        self.loader = None

        self.mainframe = ttk.Frame(self, padding="10 10 10 10")
        self.mainframe.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))

        # Input block
        self.inputFormatLabel = ttk.Label(self.mainframe, text="Input Format:")
        self.inputFormatLabel.grid(row=0, column=0, pady=5, padx=5)
        self.inputFormatCombo = ttk.Combobox(self.mainframe, values=["Excel Interlinear"])
        self.inputFormatCombo.bind('<<ComboboxSelected>>', lambda e: self.inputLoadButton.state(['!disabled']))
        self.inputFormatCombo.grid(row=0, column=1, pady=5, padx=5)
        self.inputLoadButton = ttk.Button(
            self.mainframe, text="Select input file & load",
            state='disabled', command=self.load_file_begin)
        self.inputLoadButton.grid(row=0, column=2, pady=5, padx=5)
        self.loadProgressLabel = ttk.Label(self.mainframe, text="")
        self.loadProgressLabel.grid(row=1, column=0, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.loadProgress = ttk.Progressbar(
            self.mainframe, orient=tk.HORIZONTAL, mode='determinate', maximum=1.0)
        self.loadProgress.grid(row=1, column=1, columnspan=2, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.hide_load_progress()

        # Reminder
        reminderText = ' '.join([
            "In order for data to be correctly imported into FLEx,",
            "the writing system codes in the input data must match those in your FLEx project.",
            "(You can find the writing systems in your FLEx project under Tools -> Configure -> Writing Systems...)",
            "Vernacular is also called baseline;",
            "gloss is also called word gloss or literal translation; I need to double check these."
        ])  # TODO: check labels in FLEx
        self.reminderLabel = ttk.Label(self.mainframe, text=reminderText)
        self.reminderLabel.grid(row=3, column=0, columnspan=3, pady=5, padx=5)
        self.reminderLabel.config(wraplength=400)

        # Writing systems
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
        self.update_writing_systems()

        # Output block
        self.outputFormatLabel = ttk.Label(self.mainframe, text="Output Format:")
        self.outputFormatLabel.grid(row=8, column=0, pady=5, padx=5)
        self.outputFormatCombo = ttk.Combobox(self.mainframe, values=["FlexText Interlinear"])
        self.outputFormatCombo.bind('<<ComboboxSelected>>', lambda e: self.update_convert_button_state())
        self.outputFormatCombo.grid(row=8, column=1, pady=5, padx=5)
        self.convertButton = ttk.Button(self.mainframe, text="Select output file & convert", state='disabled', command=self.convert)
        self.convertButton.grid(row=8, column=2, pady=5, padx=5)
        self.convertProgressLabel = ttk.Label(self.mainframe, text="")
        self.convertProgressLabel.grid(row=9, column=0, pady=5, padx=5)
        self.convertProgress = ttk.Progressbar(self.mainframe, orient=tk.HORIZONTAL, mode='indeterminate')
        self.convertProgress.grid(row=9, column=1, columnspan=2, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.hide_convert_progress()
        self.completeLabel = ttk.Label(self.mainframe, text="")
        self.completeLabel.grid(row=10, column=1, pady=5, padx=5)

        # Error display
        self.errorDisplay = ttk.Label(self.mainframe, text="", foreground="red", wraplength=400)
        self.errorDisplay.grid(row=11, column=0, columnspan=3, pady=5, padx=5)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.rowconfigure(1, minsize=100)
        self.rowconfigure(9, minsize=100)
        self.rowconfigure(10, minsize=100)

    def add_error_msg(self, errorString):
        """
        Add an error message to the bottom of errorDisplay.
        """

        currentText = self.errorDisplay.cget('text')
        newText = '\n'.join([currentText, errorString])
        self.errorDisplay.config(text=newText)

    def update_convert_button_state(self):
        """
        Enable the convert button if input is loaded and output format is selected.
        """

        if self.is_data_loaded and self.writing_systems_ready and self.outputFormatCombo.get():
            self.convertButton.state(['!disabled'])
        else:
            self.convertButton.state(['disabled'])

    def get_one_writing_system(self, metadataElement):
        """
        Checks and returns the text of a writing system metadata element.

        self.get_one_writing_system(metadataElement) -> displayText, isValid

        Input:
          metadataElement: the result of metadata.find('writing_system_vernacular'), etc.
        Output:
          displayText: text to display in GUI (2- or 3-letter code, or "(not found)")
          isValid: True if writing system code is valid (2 or 3 letters), False otherwise
        """

        if metadataElement is None:
            return "(not found)", False
        wsText = metadataElement.text
        if not wsText:
            return "(not found)", False
        elif not isinstance(wsText, str):
            self.add_error_msg("Error: writing system code is not a string")
            return "(invalid type)", False
        elif len(wsText) < 2 or len(wsText) > 3:
            self.add_error_msg("Error: writing system code must be 2 or 3 letters")
            if len(wsText) > 8:
                wsText = wsText[:5] + '...'
            return "Invalid code: " + wsText, False
        else:
            return wsText, True

    def update_writing_systems(self):
        if self.is_data_loaded:
            metadata = self.intermediate_xml.find('text_metadata')
            if metadata is not None:
                displayTextVernacular, isValidVernacular = self.get_one_writing_system(metadata.find('writing_system_vernacular'))
                displayTextGloss, isValidGloss = self.get_one_writing_system(metadata.find('writing_system_gloss'))
                displayTextFree, isValidFree = self.get_one_writing_system(metadata.find('writing_system_free'))
                self.wsVernacular.config(text=displayTextVernacular)
                self.wsGloss.config(text=displayTextGloss)
                self.wsFree.config(text=displayTextFree)
                self.writing_systems_ready = isValidVernacular and isValidGloss and isValidFree
                if not self.writing_systems_ready:
                    self.add_error_msg("All writing system codes must be valid in order to convert.")
            else:
                self.add_error_msg("Error: input file metadata not found")
                self.writing_systems_ready = False
                self.wsVernacular.config(text="(not found)")
                self.wsGloss.config(text="(not found)")
                self.wsFree.config(text="(not found)")
        else:
            self.writing_systems_ready = False
            self.wsVernacular.config(text="(not loaded)")
            self.wsGloss.config(text="(not loaded)")
            self.wsFree.config(text="(not loaded)")

    def hide_load_progress(self):
        self.loadProgressLabel.config(text="")
        self.loadProgress.grid_remove()

    def show_load_progress(self):
        self.loadProgressLabel.config(text="Loading...")
        self.loadProgress.grid()

    def hide_convert_progress(self):
        self.convertProgressLabel.config(text="")
        self.convertProgress.grid_remove()

    def show_convert_progress(self):
        self.convertProgressLabel.config(text="Converting...")
        self.convertProgress.grid()

    def load_file_begin(self):
        """
        Loads a file from a file dialog and processes it into intermediate XML.

        This is the first step. See also:
            load_file_next()
            load_file_success()
            load_file_end()
        """

        print('load_file_begin 1')
        self.is_data_loaded = False
        self.intermediate_xml = None
        self.inputFileName = None
        self.errorDisplay.config(text="")   # reset error messages
        self.completeLabel.config(text="")  # reset "Complete!" message
        formatString = self.inputFormatCombo.get()
        if formatString == "Excel Interlinear":
            filetypelist = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        elif formatString == "XLingPaper Interlinear":
            filetypelist = [("XLingPaper files", "*.xml"), ("All files", "*.*")]
        else:
            raise ValueError("Unsupported input format selected. Add code here")

        filepath = filedialog.askopenfilename(
            title="Load a file",
            filetypes=filetypelist)
        if not filepath:
            return None     # User cancelled
        if not os.path.exists(filepath):
            self.add_error_msg("Error: File does not exist.")
            return None

        print('load_file_begin 2')
        self.show_load_progress()
        self.loadProgress["value"] = 0.0
        if formatString == "Excel Interlinear":
            try:
                print('load_file_begin 3')
                self.loader = ExcelInterlinearLoader(filepath)
                print('load_file_begin 4')
            except Exception as e:
                self.add_error_msg(f"Error initializing ExcelInterlinearLoader:\n{traceback.format_exc()}")
                return None
        # tell tkinter to start when ready

        print('load_file_begin 5')
        self.update_idletasks()
        self.after(self.AFTER_DELAY_MS, self.load_file_next)
        print('load_file_begin 6')
        # TODO think about this more
        # TODO update requirements for InterlinearLoader class

    def load_file_next(self):
        """
        Run self.loader.next_step(), handling errors and completion events.
        """

        print('  load_file_next 1')
        if self.loader is not None:
            if self.loader.isdone:
                self.loadProgress["value"] = 1.0
                self.load_file_success()
                self.load_file_end()
            else:
                try:
                    self.loader.next_step()
                except Exception as e:
                    self.add_error_msg(traceback.format_exc(e))
                    self.load_file_end()
                else:   # if no error
                    self.loadProgress["value"] = self.loader.progress
                    # tell tkinter to do next step when ready
                    self.after(self.AFTER_DELAY_MS, self.load_file_next)
        else:
            raise RuntimeError('Cannot run next_load_step without loader')
        
    def load_file_success(self):
        """
        Display warnings and update status
        """

        if self.loader is None or self.loader.next_step is not None:
            raise Exception('Cannot show loader warnings in this state')
        error_messages = "\n".join(self.loader.warning_list)
        self.add_error_msg(error_messages)
        
        self.intermediate_xml = self.loader.xml_root
        self.is_data_loaded = True
        self.inputFileName = self.loader.loadname
        self.loadProgressLabel.config(text="Loading complete!")

    def load_file_end(self):
        """
        Finish progressbar and update status
        """
        
        # self.loadProgress.stop()
        # self.hide_load_progress()
        self.update_writing_systems()
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
        initialpath, initialname = os.path.split(self.inputFileName)
        filenamebase, _ = os.path.splitext(initialname)
        initialfilename = filenamebase + ".flextext"
        filepath = filedialog.asksaveasfilename(
            title="Save conversion output",
            initialdir=initialpath, initialfile=initialfilename,
            defaultextension=".flextext",
            filetypes=[("FlexText files", "*.flextext"), ("All files", "*.*")])
        if not filepath:
            return None

        self.show_convert_progress()
        self.convertProgress.start()
        try:
            (flextext_xml, missing_freetrans_count) = transform_to_flextext_dom(
                self.intermediate_xml, self.wsVernacular.cget('text'), self.wsGloss.cget('text'), self.wsFree.cget('text'))
        except Exception:
            self.add_error_msg(f"Conversion error:\n{traceback.format_exc()}")
            return None
        pretty_xml = self.prettify_xml(flextext_xml)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
        self.convertProgress.stop()
        self.hide_convert_progress()
        self.completeLabel.config(text="Conversion complete!")

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