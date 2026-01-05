#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, filedialog
import os
import traceback
from xml.etree.ElementTree import tostring
from xml.dom import minidom

from InterlinearLoaders import ExcelInterlinearLoader
from excel_to_xml import convert_excel_to_xml_dom
from xml_to_flextext import transform_to_flextext_dom


class Converter(tk.Tk):
    def __init__(self):
        """
        GUI application for converting interlinear data.

        __init__ method sets up the main window and all widgets.
        """

        super().__init__()
        self.AFTER_DELAY_MS = 1 # milliseconds. For allowing GUI & progressbar to update during processing
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
        self.inputFormatCombo.grid(row=0, column=1, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.inputLoadButton = ttk.Button(
            self.mainframe, text="Select input file & load",
            state='disabled', command=self.load_file_begin)
        self.inputLoadButton.grid(row=0, column=2, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.loadProgressLabel = ttk.Label(self.mainframe, text="")
        self.loadProgressLabel.grid(row=1, column=0, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.loadProgress = ttk.Progressbar(
            self.mainframe, orient=tk.HORIZONTAL, mode='determinate', maximum=1.0)
        self.loadProgress.grid(row=1, column=1, columnspan=2, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.hide_load_progress()

        # Reminder
        reminderText = ' '.join([
            "Note:",
            "\n\nIn order for data to be correctly imported into FLEx,",
            "the writing system codes in the input data must match those in your FLEx project.",
            "\n\nSpecifically, go in the menu to Tools -> Configure -> Writing Systems...",
            "to set the vernacular writing system (baseline),",
            "and the writing system(s) available for analysis (including gloss and free translation).",
            "In addition, open the 'Gloss' or 'Analyze' tab of your text then go in the menu to",
            "Tools -> Configure -> Interlinear... to set the writing systems",
            "for the Word Gloss and Free Translation."
        ])  # TODO: check newer FLEx versions for updated menu paths
        self.reminderLabel = ttk.Label(self.mainframe, text=reminderText)
        self.reminderLabel.grid(row=3, column=0, columnspan=3, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.reminderLabel.config(wraplength=1)  # Initial value, updated to match width of window
        def update_wraplength(event):
            self.reminderLabel.config(wraplength=event.width)
        self.reminderLabel.bind('<Configure>', update_wraplength)

        # Writing systems frame
        self.wsFrame = ttk.Frame(self.mainframe)
        self.wsFrame.grid(row=4, column=0, columnspan=3, pady=1, padx=0, sticky=(tk.W, tk.E))
        self.wsVernacularLabel = ttk.Label(self.wsFrame, text="\tVernacular writing system:", anchor='w', justify='left')
        self.wsVernacularLabel.grid(row=0, column=0, pady=1, padx=5, sticky='w')
        self.wsVernacular = ttk.Label(self.wsFrame, text="(not loaded)", anchor='w')
        self.wsVernacular.grid(row=0, column=1, pady=1, padx=5)
        self.wsGlossLabel = ttk.Label(self.wsFrame, text="\tGloss writing system:", anchor='w', justify='left')
        self.wsGlossLabel.grid(row=1, column=0, pady=1, padx=5, sticky='w')
        self.wsGloss = ttk.Label(self.wsFrame, text="(not loaded)", anchor='w')
        self.wsGloss.grid(row=1, column=1, pady=1, padx=5)
        self.wsFreeLabel = ttk.Label(self.wsFrame, text="\tFree trans. writing system:", anchor='w', justify='left')
        self.wsFreeLabel.grid(row=2, column=0, pady=1, padx=5, sticky='w')
        self.wsFree = ttk.Label(self.wsFrame, text="(not loaded)", anchor='w')
        self.wsFree.grid(row=2, column=1, pady=1, padx=5)
        self.extraSpace = ttk.Label(self.wsFrame, text="\n")
        self.extraSpace.grid(row=3, column=1)
        self.update_writing_systems()

        # Output block
        self.outputFormatLabel = ttk.Label(self.mainframe, text="Output Format:")
        self.outputFormatLabel.grid(row=8, column=0, pady=5, padx=5)
        self.outputFormatCombo = ttk.Combobox(self.mainframe, values=["FlexText Interlinear"])
        self.outputFormatCombo.bind('<<ComboboxSelected>>', lambda e: self.update_convert_button_state())
        self.outputFormatCombo.grid(row=8, column=1, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.convertButton = ttk.Button(self.mainframe, text="Select output file & convert", state='disabled', command=self.convert)
        self.convertButton.grid(row=8, column=2, pady=5, padx=5)
        self.convertProgressLabel = ttk.Label(self.mainframe, text="")
        self.convertProgressLabel.grid(row=9, column=0, pady=5, padx=5)
        self.convertProgress = ttk.Progressbar(self.mainframe, orient=tk.HORIZONTAL, mode='determinate', maximum=1.0)
        self.convertProgress.grid(row=9, column=1, columnspan=2, pady=5, padx=5, sticky=(tk.W, tk.E))
        self.hide_convert_progress()

        # Error display
        default_font = ttk.Style().lookup('TLabel', 'font') # because the tk.Text widget has a different default
        self.errorDisplay = tk.Text(
            self.mainframe, wrap='word', height=9, width=50, state='disabled',
            borderwidth=2, relief='sunken', font=default_font,
            yscrollcommand=lambda *args: self.errorDisplayScrollbar.set(*args))
        self.errorDisplay.grid(row=11, column=0, columnspan=3, pady=5, padx=(5,0), sticky=(tk.N, tk.S, tk.W, tk.E))
        self.errorDisplayScrollbar = ttk.Scrollbar(self.mainframe, orient='vertical', command=self.errorDisplay.yview)
        self.errorDisplayScrollbar.grid(row=11, column=3, sticky=(tk.N, tk.S, tk.W))

        # Make mainframe expand with main window
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        # Make center column expand horizontally with main window
        self.mainframe.columnconfigure(1, weight=1)
        # Make error message box expand vertically with main window
        self.mainframe.rowconfigure(11, weight=1)

    def add_error_msg(self, errorString):
        """
        Add an error message to the bottom of errorDisplay.
        """

        # TODO Add color-coding for errors vs. warnings vs. info. https://tkdocs.com/tutorial/text.html
        self.errorDisplay.config(state='normal')
        see_this = self.errorDisplay.index('end')   # Position at end of current text
        self.errorDisplay.insert('end', '\n' + errorString) # Start a new line
        self.errorDisplay.see(see_this) # Make sure the end of previous text is visible.
        # So the top of a big block of load warnings will be visible first, instead of the end of the block.
        # But after that, if a new message is printed, the text box scrolls so it is visible.
        # This behavior makes the most sense to me...
        self.errorDisplay.config(state='disabled')

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

        If writing system code is not valid, displayText indicates the problem.

        Input:
          metadataElement: the result of metadata.find('writing_system_vernacular'), etc.
        Output:
          displayText: text to display in GUI (code string, or "(not found)")
          isValid: True if writing system code is valid (non-empty string), False otherwise
        """

        if metadataElement is None:
            return "(not found)", False
        wsText = metadataElement.text
        if not wsText:
            return "(not found)", False
        elif not isinstance(wsText, str):
            self.add_error_msg("❌ Error: writing system code is not a string")
            return "(invalid type)", False
        else:
            return wsText, True

    def update_writing_systems(self):
        """
        Update the writing system display fields based on data and status.

        Also set self.writing_systems_ready.
        """

        self.writing_systems_ready = False
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
                    self.add_error_msg("❌ All writing system codes must be valid in order to convert.")
            else:
                self.add_error_msg("❌ Error: input file metadata not found")
                self.wsVernacular.config(text="(not found)")
                self.wsGloss.config(text="(not found)")
                self.wsFree.config(text="(not found)")
        else:
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

        This is the first step, which includes the file selection dialog
        and initializing the Loader object.

        See also:
            load_file_next()
            load_file_success()
            load_file_end()
        """

        self.is_data_loaded = False
        self.intermediate_xml = None
        self.inputFileName = None
        # delete error messages
        self.errorDisplay.config(state='normal')
        self.errorDisplay.delete('1.0', 'end')
        self.errorDisplay.config(state='disabled')

        # Check input type from dropdown
        formatString = self.inputFormatCombo.get()
        if formatString == "Excel Interlinear":
            filetypelist = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        else:
            # User should never get this error, because button is disabled until input type is selected.
            # Only if inputFormatCombo is updated but this code is not updated.
            raise ValueError("Unsupported input format selected. Add code here")

        # Get a filename using a dialog box
        self.inputFileName = filedialog.askopenfilename(title="Load a file", filetypes=filetypelist)
        if not self.inputFileName:
            return None     # User cancelled
        if not os.path.exists(self.inputFileName):
            self.add_error_msg("❌ Error: File does not exist.")
            return None

        # Initialize progressbar
        self.show_load_progress()
        self.loadProgress["value"] = 0.0

        # Initialize Loader object
        if formatString == "Excel Interlinear":
            try:
                self.loader = ExcelInterlinearLoader(self.inputFileName)
            except Exception as e:
                self.add_error_msg(f"❌ Error initializing ExcelInterlinearLoader:\n{traceback.format_exc()}")
                return None

        # tkinter is single-threaded. So we schedule each incremental step of processing
        #   to give tkinter time to refresh the GUI, including the progressbar,
        #   so it doesn't appear frozen.
        self.update_idletasks() # needed to draw progressbar
        self.after(self.AFTER_DELAY_MS, self.load_file_next)

    def load_file_next(self):
        """
        Run the next processing step of the loader, handling errors and completion events.
        """

        if self.loader is not None:
            if self.loader.isdone:
                # Finalize progressbar, update statuses
                self.loadProgress["value"] = 1.0
                self.load_file_success()
                self.update_writing_systems()
                self.update_convert_button_state()
            else:
                try:
                    self.loader.next_step()
                    # The Loader object handles what the next step is
                except Exception as e:
                    self.add_error_msg('❌ Loading error: ' + traceback.format_exc())
                    # Update statuses
                    self.update_writing_systems()
                    self.update_convert_button_state()
                else:   # if no error
                    self.loadProgress["value"] = self.loader.progress
                    # schedule another step with tkinter
                    self.after(self.AFTER_DELAY_MS, self.load_file_next)
        else:
            # should not be possible for user to get this error
            raise RuntimeError('Cannot run next_load_step without loader')
        
    def load_file_success(self):
        """
        Display load warnings and update status after successful loading
        """

        if not self.loader.isdone:  # should not be possible for user to get this error
            raise Exception('Cannot show loader warnings when loader is not done')
        if self.loader.warning_list:
            error_messages = '⚠️' + '\n⚠️  '.join(self.loader.warning_list)
            self.add_error_msg(error_messages)

        # Get XML data and update status
        self.intermediate_xml = self.loader.xml_root
        self.is_data_loaded = True
        self.loadProgressLabel.config(text="Loading complete!")

    def convert(self):
        """
        Sets an output file from a file dialog and converts into target format.
        """

        # TODO: Refactor conversion code into a Converter or Exporter class, similar to Loader.
        #   If the class is set up with next_step() like Loader, it would allow
        #   the progressbar to work (although so far, this step is very quick).
        #   Whether it is designed with next_step() or not, a class structure would also
        #   help ensure that the GUI can use the same interface with different export formats.

        # Check output format type from dropdown
        formatString = self.outputFormatCombo.get()
        if formatString == "FlexText Interlinear":
            filetypelist = [("FlexText files", "*.flextext"), ("All files", "*.*")]
        else:
            # User should never get this error
            raise ValueError("Unsupported output format selected. Add code here")

        # Construct suggested filename
        initialpath, initialname = os.path.split(self.inputFileName)
        filenamebase, _ = os.path.splitext(initialname)
        initialfilename = filenamebase + ".flextext"
        # Provide dialog for user to confirm suggested filename, or rename
        filepath = filedialog.asksaveasfilename(
            title="Save conversion output",
            initialdir=initialpath, initialfile=initialfilename,
            defaultextension=".flextext",
            filetypes=[("FlexText files", "*.flextext"), ("All files", "*.*")])
        if not filepath:
            return None

        # Initialize progressbar
        self.show_convert_progress()
        self.convertProgress["value"] = 0.0
        self.update_idletasks() # let GUI update to show progressbar

        try:
            # Perform conversion to output format
            (flextext_xml, missing_freetrans_count) = transform_to_flextext_dom(
                self.intermediate_xml, self.wsVernacular.cget('text'), self.wsGloss.cget('text'), self.wsFree.cget('text'))
            # TODO exporter could read writing system codes from metadata, instead of as input args
        except Exception:
            self.add_error_msg(f"❌ Conversion error:\n{traceback.format_exc()}")
            return None

        # Make the output XML pretty. This should actually go in exporter code
        pretty_xml = self.prettify_xml(flextext_xml)

        # Write to file
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)

        # Finalize progressbar, etc.
        self.convertProgress["value"] = 1.0
        self.hide_convert_progress()
        self.convertProgressLabel.config(text="Conversion complete!")
        self.add_error_msg(f"\nWritten to file at {filepath}") # extra blank line

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