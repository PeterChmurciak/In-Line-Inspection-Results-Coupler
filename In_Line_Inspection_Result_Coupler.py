# -*- coding: utf-8 -*-
"""
In-Line Inspection Result Coupler app.
Copyright (C) 2022  Peter Chmurčiak

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see https://www.gnu.org/licenses/.
"""
import tkinter as tk  # Providing constants and running the app
from tkinter.filedialog import askopenfilename, asksaveasfilename # Managing file choice
from tkinter.messagebox import showerror, showwarning  # Interacting with user
import os  # Working with paths
import re  # Working with regular expressions to validate input
import pandas as pd  # Providing main functionality through dataframe slicing
import numpy as np  # Math constants and array friendly functions
import pdb  # Debugging

from Dependencies.Abstract_Class_Implementing_Coupler_GUI import In_Line_Inspection_Result_Coupler_GUI


class In_Line_Inspection_Result_Coupler(In_Line_Inspection_Result_Coupler_GUI):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master

    def button_pipetally_1_pressed(self):
        self.pipetally_1_path = askopenfilename(
            filetypes=(("Excel files (*.xlsx)", ("*.xls", "*.xlsx")),)
        )
        extracted_filename = os.path.basename(self.pipetally_1_path)
        self.entry_pipetally_1_filename_stringvar.set(extracted_filename)

    def button_pipetally_2_pressed(self):
        self.pipetally_2_path = askopenfilename(
            filetypes=(("Excel files (*.xlsx)", ("*.xls", "*.xlsx")),)
        )
        extracted_filename = os.path.basename(self.pipetally_2_path)
        self.entry_pipetally_2_filename_stringvar.set(extracted_filename)

    def button_output_pressed(self):
        self.output_path = asksaveasfilename(
            filetypes=(("Excel files (*.xlsx)", ("*.xls", "*.xlsx")),),
            defaultextension=".xlsx",
        )
        extracted_filename = os.path.basename(self.output_path)
        self.entry_output_filename_stringvar.set(extracted_filename)

    def entry_range_validation(self, P):
        acceptable_pattern = "^$|[A-Z]|[A-Z][1-9][0-9]{0,2}|[A-Z][1-9][0-9]{0,2}:|[A-Z][1-9][0-9]{0,2}:[A-Z]|[A-Z][1-9][0-9]{0,2}:[A-Z][1-9][0-9]{0,8}"
        found_match = re.fullmatch(acceptable_pattern, P)
        if found_match:
            return True
        else:
            return False

    def reset_setup_frame(self):
        self.dictionary_of_steps = {"Krok číslo 1": self.template_dictionary.copy()}
        self.listbox_list_of_steps.delete(0, tk.END)
        self.listbox_list_of_steps.insert(0, *self.dictionary_of_steps.keys())
        self.listbox_list_of_steps.selection_set(0)
        self.button_load_step_pressed()

    def button_confirm_pressed(self):
        self.reset_setup_frame()
        paths_and_ranges_are_valid = True

        if not self.pipetally_1_path:
            showerror("Chyba", "Zadaná cesta k súboru PipeTally 1 je neplatná!")
            paths_and_ranges_are_valid = False
        if not self.pipetally_2_path:
            showerror("Chyba", "Zadaná cesta k súboru PipeTally 2 je neplatná!")
            paths_and_ranges_are_valid = False
        if not self.output_path:
            showerror("Chyba", "Zadaná cesta k výstupnému súboru je neplatná!")
            paths_and_ranges_are_valid = False
        matched_range_1 = re.fullmatch(
            "([A-Z])([1-9][0-9]*):([A-Z])([1-9][0-9]*)",
            self.entry_pipetally_1_range.get(),
        )
        if matched_range_1:
            (
                self.start_column_1,
                self.start_row_1,
                self.end_column_1,
                self.end_row_1,
            ) = matched_range_1.groups()
            self.start_row_1 = int(self.start_row_1)
            self.end_row_1 = int(self.end_row_1)
            if ord(self.start_column_1) >= ord(self.end_column_1) or (
                self.start_row_1 >= self.end_row_1
            ):
                showerror(
                    "Chyba",
                    "Zadaný rozsah tabuľky pre súbor PipeTally 1 je neplatný. Začiatočný stĺpec nesmie predbiehať konečný stĺpec rozsahu a tak isto začiatočný riadok rozsahu nemôže byť väčší ako konečný riadok rozsahu. Taktiež stĺpce/riadky začiatku/konca si nesmú byť rovné.",
                )
                paths_and_ranges_are_valid = False
            else:
                self.valid_pipetally_1_columns = self.ALPHABET[
                    self.ALPHABET.index(self.start_column_1) : self.ALPHABET.index(
                        self.end_column_1
                    )
                    + 1
                ]
                optionmenu_stringvar_changed = False
                for optionmenu, optionmenu_stringvar in zip(
                    self.optionmenus_pipetally_1, self.optionmenu_stringvars_pipetally_1
                ):
                    optionmenu["menu"].delete(0, tk.END)
                    for valid_column in ["", *self.valid_pipetally_1_columns]:
                        optionmenu["menu"].add_command(
                            label=valid_column,
                            command=tk._setit(optionmenu_stringvar, valid_column),
                        )
                    if (
                        optionmenu_stringvar.get() in self.ALPHABET
                        and optionmenu_stringvar.get()
                        not in self.valid_pipetally_1_columns
                    ):

                        def letter_difference_function(letter):
                            return abs(ord(letter) - ord(optionmenu_stringvar.get()))

                        closest_letter_by_value = min(
                            self.valid_pipetally_1_columns,
                            key=letter_difference_function,
                        )
                        optionmenu_stringvar.set(closest_letter_by_value)
                        optionmenu_stringvar_changed = True
                if optionmenu_stringvar_changed:
                    showwarning(
                        "Upozornenie",
                        "Hodnoty niektorých stĺpcov pre PipeTally 1 boli zmenené tak, aby spadali do uvedeného rozsahu tabuľky. Uistite sa že všetky používané stĺpce sú aj po zmene definované správne.",
                    )
                    paths_and_ranges_are_valid = False
        else:
            showerror(
                "Chyba",
                "Zadaný rozsah tabuľky pre súbor PipeTally 1 je neplatný. Uistite sa že rozsah je v správnom formáte používanom v prostredí Excel (e.g. A10:E100).",
            )
            paths_and_ranges_are_valid = False
        matched_range_2 = re.fullmatch(
            "([A-Z])([1-9][0-9]*):([A-Z])([1-9][0-9]*)",
            self.entry_pipetally_2_range.get(),
        )
        if matched_range_2:
            (
                self.start_column_2,
                self.start_row_2,
                self.end_column_2,
                self.end_row_2,
            ) = matched_range_2.groups()
            self.start_row_2 = int(self.start_row_2)
            self.end_row_2 = int(self.end_row_2)
            if ord(self.start_column_2) >= ord(self.end_column_2) or (
                self.start_row_2 >= self.end_row_2
            ):
                showerror(
                    "Chyba",
                    "Zadaný rozsah tabuľky pre súbor PipeTally 2 je neplatný. Začiatočný stĺpec nesmie predbiehať konečný stĺpec rozsahu a tak isto začiatočný riadok rozsahu nemôže byť väčší ako konečný riadok rozsahu. Taktiež stĺpce/riadky začiatku/konca si nesmú byť rovné.",
                )
                paths_and_ranges_are_valid = False
            else:
                self.valid_pipetally_2_columns = self.ALPHABET[
                    self.ALPHABET.index(self.start_column_2) : self.ALPHABET.index(
                        self.end_column_2
                    )
                    + 1
                ]
                optionmenu_stringvar_changed = False
                for optionmenu, optionmenu_stringvar in zip(
                    self.optionmenus_pipetally_2, self.optionmenu_stringvars_pipetally_2
                ):
                    optionmenu["menu"].delete(0, tk.END)
                    for valid_column in ["", *self.valid_pipetally_2_columns]:
                        optionmenu["menu"].add_command(
                            label=valid_column,
                            command=tk._setit(optionmenu_stringvar, valid_column),
                        )
                    if (
                        optionmenu_stringvar.get() in self.ALPHABET
                        and optionmenu_stringvar.get()
                        not in self.valid_pipetally_2_columns
                    ):

                        def letter_difference_function(letter):
                            return abs(ord(letter) - ord(optionmenu_stringvar.get()))

                        closest_letter_by_value = min(
                            self.valid_pipetally_2_columns,
                            key=letter_difference_function,
                        )
                        optionmenu_stringvar.set(closest_letter_by_value)
                        optionmenu_stringvar_changed = True
                if optionmenu_stringvar_changed:
                    showwarning(
                        "Upozornenie",
                        "Hodnoty niektorých stĺpcov pre PipeTally 2 boli zmenené tak, aby spadali do uvedeného rozsahu tabuľky. Uistite sa že všetky používané stĺpce sú aj po zmene definované správne.",
                    )
                    paths_and_ranges_are_valid = False
        else:
            showerror(
                "Chyba",
                "Zadaný rozsah tabuľky pre súbor PipeTally 2 je neplatný. Uistite sa že rozsah je v správnom formáte používanom v prostredí Excel (e.g. A10:E100).",
            )
            paths_and_ranges_are_valid = False
        self.specified_column_pairs = [False] * len(
            self.optionmenu_stringvars_pipetally_1
        )
        if paths_and_ranges_are_valid:
            for index, (optionmenu_stringvar_1, optionmenu_stringvar_2) in enumerate(
                zip(
                    self.optionmenu_stringvars_pipetally_1,
                    self.optionmenu_stringvars_pipetally_2,
                )
            ):
                if (
                    (optionmenu_stringvar_1.get() in self.valid_pipetally_1_columns)
                    and (optionmenu_stringvar_1.get() != "")
                    and (optionmenu_stringvar_2.get() in self.valid_pipetally_2_columns)
                    and (optionmenu_stringvar_2.get() != "")
                ):
                    self.specified_column_pairs[index] = True
        checkbutton_states = [
            *self.specified_column_pairs[:-2],
            self.specified_column_pairs[-1] and self.specified_column_pairs[-2],
        ]
        for enabled_state, checkbutton in zip(checkbutton_states, self.checkbuttons):
            if enabled_state:
                checkbutton.configure(state="normal")
            else:
                checkbutton.configure(state="disabled")
        # pdb.set_trace()

        if paths_and_ranges_are_valid and not any(self.specified_column_pairs):
            showerror(
                "Chyba",
                "Žiadna súvisiaca dvojica stĺpcov pre PipeTally 1 a PipeTally 2 nebola špecifikovaná. Párovanie údajov (riadkov) je možné iba keď stĺpec daného párovacieho kritéria je známy pre obe tabuľky.",
            )
        elif paths_and_ranges_are_valid and any(self.specified_column_pairs):

            self.pipetally_1_data = pd.read_excel(
                self.pipetally_1_path,
                header=self.start_row_1 - 1,
                usecols=self.start_column_1 + ":" + self.end_column_1,
                nrows=self.end_row_1 - self.start_row_1 + 1,
                na_values=["", " "],
            )
            self.column_letter_to_name_dictionary_1 = dict(
                zip(
                    self.valid_pipetally_1_columns, self.pipetally_1_data.columns.values
                )
            )

            self.pipetally_2_data = pd.read_excel(
                self.pipetally_2_path,
                header=self.start_row_2 - 1,
                usecols=self.start_column_2 + ":" + self.end_column_2,
                nrows=self.end_row_2 - self.start_row_2 + 1,
                na_values=["", " "],
            )
            self.column_letter_to_name_dictionary_2 = dict(
                zip(
                    self.valid_pipetally_2_columns, self.pipetally_2_data.columns.values
                )
            )

            try:
                self.listbox_description_1_data = []
                self.listbox_description_2_data = []
                if self.specified_column_pairs[1]:
                    self.listbox_description_1_data = (
                        self.pipetally_1_data[
                            self.column_letter_to_name_dictionary_1[
                                self.optionmenu_description_column_1_stringvar.get()
                            ]
                        ]
                        .sort_values()
                        .unique()
                    )
                    self.listbox_description_2_data = (
                        self.pipetally_2_data[
                            self.column_letter_to_name_dictionary_2[
                                self.optionmenu_description_column_2_stringvar.get()
                            ]
                        ]
                        .sort_values()
                        .unique()
                    )
            except TypeError:
                showerror(
                    "Chyba",
                    'Stĺpce špecifikované pre kategóriu "Description" sa javia ako nekompatibilné. Uistite sa že boli zvolené správne stĺpce.',
                )
            try:
                self.listbox_wall_side_1_data = []
                self.listbox_wall_side_2_data = []
                if self.specified_column_pairs[7]:
                    self.listbox_wall_side_1_data = (
                        self.pipetally_1_data[
                            self.column_letter_to_name_dictionary_1[
                                self.optionmenu_wall_side_column_1_stringvar.get()
                            ]
                        ]
                        .sort_values()
                        .unique()
                    )
                    self.listbox_wall_side_2_data = (
                        self.pipetally_2_data[
                            self.column_letter_to_name_dictionary_2[
                                self.optionmenu_wall_side_column_2_stringvar.get()
                            ]
                        ]
                        .sort_values()
                        .unique()
                    )
            except TypeError:
                showerror(
                    "Chyba",
                    'Stĺpce špecifikované pre kategóriu "Wall Side" sa javia ako nekompatibilné. Uistite sa že boli zvolené správne stĺpce.',
                )
            # self.frame_setup_and_run.pack(expand='true', fill='both', padx='5', pady='0 5', side='top')

    def entry_number_validation(self, P):
        acceptable_pattern = "^$|0|[1-9][0-9]{0,3}"
        found_match = re.fullmatch(acceptable_pattern, P)
        if found_match:
            return True
        else:
            return False

    def entry_orientation_validation(self, P):
        acceptable_pattern = "^$|([0-9][0-9]?):?|([0-9][0-9]?):([0-9][0-9]?)"
        found_match = re.fullmatch(acceptable_pattern, P)
        if found_match:
            hh1, hh2, mm2 = found_match.groups()
            if hh1:
                return int(hh1) <= 12
            elif hh2 and mm2:
                if int(hh2) == 12:
                    return int(mm2) == 0
                else:
                    return int(mm2) <= 59
            else:
                return True
        else:
            return False

    def button_add_step_pressed(self):
        index_to_insert_the_new_step = self.listbox_list_of_steps.curselection()[0]
        new_dictionary = {}
        index = 0
        for value in self.dictionary_of_steps.values():
            new_step_name = f"Krok číslo {index + 1}"
            new_dictionary[new_step_name] = value
            if index == index_to_insert_the_new_step:
                index = index + 1
                new_step_name = f"Krok číslo {index + 1}"
                new_dictionary[new_step_name] = self.template_dictionary.copy()
            index = index + 1
        self.dictionary_of_steps = new_dictionary
        self.listbox_list_of_steps.delete(0, tk.END)
        self.listbox_list_of_steps.insert(0, *self.dictionary_of_steps.keys())
        self.listbox_list_of_steps.selection_set(index_to_insert_the_new_step + 1)
        self.button_load_step_pressed()

    def button_save_step_pressed(self):
        step_name = self.listbox_list_of_steps.get(
            self.listbox_list_of_steps.curselection()
        )

        self.dictionary_of_steps[step_name]["log_distance"] = (
            self.checkbutton_log_dist_intvar.get(),
            self.entry_log_dist_tolerance.get(),
        )
        if (
            self.checkbutton_log_dist_intvar.get()
            and self.entry_log_dist_tolerance.get() == ""
        ):
            showwarning("Upozornenie", 'Pole "Log Distance" je prázdne!')
        self.dictionary_of_steps[step_name]["wt"] = (
            self.checkbutton_wt_intvar.get(),
            self.entry_wt_tolerance.get(),
        )
        if self.checkbutton_wt_intvar.get() and self.entry_wt_tolerance.get() == "":
            showwarning("Upozornenie", 'Pole "WT" je prázdne!')
        self.dictionary_of_steps[step_name]["depth"] = (
            self.checkbutton_depth_intvar.get(),
            self.entry_depth_tolerance.get(),
        )
        if (
            self.checkbutton_depth_intvar.get()
            and self.entry_depth_tolerance.get() == ""
        ):
            showwarning("Upozornenie", 'Pole "Depth" je prázdne!')
        self.dictionary_of_steps[step_name]["length"] = (
            self.checkbutton_length_intvar.get(),
            self.entry_length_tolerance.get(),
        )
        if (
            self.checkbutton_length_intvar.get()
            and self.entry_length_tolerance.get() == ""
        ):
            showwarning("Upozornenie", 'Pole "Length" je prázdne!')
        self.dictionary_of_steps[step_name]["width"] = (
            self.checkbutton_width_intvar.get(),
            self.entry_width_tolerance.get(),
        )
        if (
            self.checkbutton_width_intvar.get()
            and self.entry_width_tolerance.get() == ""
        ):
            showwarning("Upozornenie", 'Pole "Width" je prázdne!')
        self.dictionary_of_steps[step_name]["orientation"] = (
            self.checkbutton_orientation_intvar.get(),
            self.entry_orientation_tolerance.get(),
        )
        if (
            self.checkbutton_orientation_intvar.get()
            and self.entry_orientation_tolerance.get() == ""
        ):
            showwarning("Upozornenie", 'Pole "Orientation" je prázdne!')
        self.dictionary_of_steps[step_name]["coordinate_distance"] = (
            self.checkbutton_coordinate_distance_intvar.get(),
            self.entry_coordinate_distance_tolerance.get(),
        )
        if (
            self.checkbutton_coordinate_distance_intvar.get()
            and self.entry_coordinate_distance_tolerance.get() == ""
        ):
            showwarning("Upozornenie", 'Pole "Coordinate Distance" je prázdne!')
        if (
            self.listbox_description_1.curselection()
            and self.listbox_description_2.curselection()
        ):
            self.dictionary_of_steps[step_name]["description"] = (
                self.checkbutton_description_intvar.get(),
                self.listbox_description_1.get(0, tk.END),
                self.listbox_description_1.curselection(),
                self.listbox_description_2.get(0, tk.END),
                self.listbox_description_2.curselection(),
            )
        elif self.checkbutton_description_intvar.get():
            showwarning(
                "Upozornenie", 'Jedno z polí pre kritérium "Description" je prázdne!'
            )
        if (
            self.listbox_wall_side_1.curselection()
            and self.listbox_wall_side_2.curselection()
        ):
            self.dictionary_of_steps[step_name]["wall_side"] = (
                self.checkbutton_wall_side_intvar.get(),
                self.listbox_wall_side_1.get(0, tk.END),
                self.listbox_wall_side_1.curselection(),
                self.listbox_wall_side_2.get(0, tk.END),
                self.listbox_wall_side_2.curselection(),
            )
        elif self.checkbutton_wall_side_intvar.get():
            showwarning(
                "Upozornenie", 'Jedno z polí pre kritérium "Wall Side" je prázdne!'
            )

    def button_load_step_pressed(self, event=None):
        step_name = self.listbox_list_of_steps.get(
            self.listbox_list_of_steps.curselection()
        )

        state, value = self.dictionary_of_steps[step_name]["log_distance"]
        self.checkbutton_log_dist_intvar.set(state)
        self.change_entry_state(
            self.entry_log_dist_tolerance, self.checkbutton_log_dist_intvar, value
        )

        state, value = self.dictionary_of_steps[step_name]["wt"]
        self.checkbutton_wt_intvar.set(state)
        self.change_entry_state(
            self.entry_wt_tolerance, self.checkbutton_wt_intvar, value
        )

        state, value = self.dictionary_of_steps[step_name]["depth"]
        self.checkbutton_depth_intvar.set(state)
        self.change_entry_state(
            self.entry_depth_tolerance, self.checkbutton_depth_intvar, value
        )

        state, value = self.dictionary_of_steps[step_name]["length"]
        self.checkbutton_length_intvar.set(state)
        self.change_entry_state(
            self.entry_length_tolerance, self.checkbutton_length_intvar, value
        )

        state, value = self.dictionary_of_steps[step_name]["width"]
        self.checkbutton_width_intvar.set(state)
        self.change_entry_state(
            self.entry_width_tolerance, self.checkbutton_width_intvar, value
        )

        state, value = self.dictionary_of_steps[step_name]["orientation"]
        self.checkbutton_orientation_intvar.set(state)
        self.change_entry_state(
            self.entry_orientation_tolerance, self.checkbutton_orientation_intvar, value
        )

        state, value = self.dictionary_of_steps[step_name]["coordinate_distance"]
        self.checkbutton_coordinate_distance_intvar.set(state)
        self.change_entry_state(
            self.entry_coordinate_distance_tolerance,
            self.checkbutton_coordinate_distance_intvar,
            value,
        )

        state, values_1, indexes_1, values_2, indexes_2 = self.dictionary_of_steps[
            step_name
        ]["description"]
        self.checkbutton_description_intvar.set(state)
        self.change_listbox_state(
            self.listbox_description_1,
            self.listbox_description_2,
            self.checkbutton_description_intvar,
            values_1,
            indexes_1,
            values_2,
            indexes_2,
        )

        state, values_1, indexes_1, values_2, indexes_2 = self.dictionary_of_steps[
            step_name
        ]["wall_side"]
        self.checkbutton_wall_side_intvar.set(state)
        self.change_listbox_state(
            self.listbox_wall_side_1,
            self.listbox_wall_side_2,
            self.checkbutton_wall_side_intvar,
            values_1,
            indexes_1,
            values_2,
            indexes_2,
        )

        self.focus_set()

    def button_delete_step_pressed(self):
        if len(self.dictionary_of_steps) > 1:
            index_of_step_to_delete = self.listbox_list_of_steps.curselection()[0]
            new_dictionary = {}
            index = 0
            for i, value in enumerate(self.dictionary_of_steps.values()):
                if i == index_of_step_to_delete:
                    continue
                new_step_name = f"Krok číslo {index + 1}"
                new_dictionary[new_step_name] = value
                index = index + 1
            self.dictionary_of_steps = new_dictionary
            self.listbox_list_of_steps.delete(0, tk.END)
            self.listbox_list_of_steps.insert(0, *self.dictionary_of_steps.keys())
            if index_of_step_to_delete in range(len(self.dictionary_of_steps)):
                selection_index = index_of_step_to_delete
            else:
                selection_index = index_of_step_to_delete - 1
            self.listbox_list_of_steps.selection_set(selection_index)
            self.button_load_step_pressed()

    def change_entry_state(self, entry, checkbutton_intvar, custom_value=""):
        if checkbutton_intvar.get():
            entry.configure(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, custom_value)
        else:
            entry.delete(0, tk.END)
            entry.configure(state="disabled")

    def change_listbox_state(
        self,
        listbox_1,
        listbox_2,
        checkbutton_intvar,
        listbox_1_values=(),
        listbox_1_indexes=(),
        listbox_2_values=(),
        listbox_2_indexes=(),
    ):
        if checkbutton_intvar.get():
            listbox_1.configure(state="normal")
            listbox_2.configure(state="normal")
            listbox_1.delete(0, tk.END)
            listbox_2.delete(0, tk.END)
            listbox_1.insert(0, *listbox_1_values)
            listbox_2.insert(0, *listbox_2_values)
            if listbox_1_indexes and listbox_2_indexes:
                for index in listbox_1_indexes:
                    listbox_1.selection_set(index)
                for index in listbox_2_indexes:
                    listbox_2.selection_set(index)
        else:
            self.frame_categoric_criterion.focus_set()
            listbox_1.delete(0, tk.END)
            listbox_2.delete(0, tk.END)
            listbox_1.configure(state="disabled")
            listbox_2.configure(state="disabled")

    def button_start_pairing_pressed(self):

        pipetally_2_paired = pd.DataFrame(
            data=None, columns=self.pipetally_2_data.columns
        )
        pipetally_2_empty_row = (
            pd.DataFrame().reindex_like(self.pipetally_2_data).iloc[0]
        )

        shrinking_Data_2 = self.pipetally_2_data

        pipetally_1_paired = pd.DataFrame(
            data=None, columns=self.pipetally_1_data.columns
        )
        pipetally_2_paired = pd.DataFrame(
            data=None, columns=self.pipetally_2_data.columns
        )
        pipetally_2_empty_row = (
            pd.DataFrame().reindex_like(self.pipetally_2_data).iloc[0]
        )

        for planned_step in self.dictionary_of_steps.values():

            if planned_step["orientation"][0]:
                if not "orientation_in_seconds" in self.pipetally_1_data.columns:
                    self.pipetally_1_data[
                        "orientation_in_seconds"
                    ] = self.pipetally_1_data[
                        self.column_letter_to_name_dictionary_1[
                            self.optionmenu_orientation_column_1_stringvar.get()
                        ]
                    ].apply(
                        self.convert_time_to_seconds
                    )
                if not "orientation_in_seconds" in self.pipetally_2_data.columns:
                    self.pipetally_2_data[
                        "orientation_in_seconds"
                    ] = self.pipetally_2_data[
                        self.column_letter_to_name_dictionary_2[
                            self.optionmenu_orientation_column_2_stringvar.get()
                        ]
                    ].apply(
                        self.convert_time_to_seconds
                    )
                    shrinking_Data_2 = self.pipetally_2_data
            for index_1, row_1 in self.pipetally_1_data.iterrows():
                working_Data_2 = shrinking_Data_2

                if planned_step["description"][0]:
                    allowed_values_1 = list(
                        np.array(planned_step["description"][1])[
                            list(planned_step["description"][2])
                        ]
                    )
                    allowed_values_2 = list(
                        np.array(planned_step["description"][3])[
                            list(planned_step["description"][4])
                        ]
                    )
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_description_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_description_column_2_stringvar.get()
                    ]
                    if not row_1[relevant_column_1] in allowed_values_1:
                        pipetally_2_paired.loc[
                            pipetally_2_paired.shape[0]
                        ] = pipetally_2_empty_row
                        continue
                    working_Data_2 = working_Data_2[
                        working_Data_2[relevant_column_2].isin(allowed_values_2)
                    ]
                if planned_step["wall_side"][0]:
                    allowed_values_1 = list(
                        np.array(planned_step["wall_side"][1])[
                            list(planned_step["wall_side"][2])
                        ]
                    )
                    allowed_values_2 = list(
                        np.array(planned_step["wall_side"][3])[
                            list(planned_step["wall_side"][4])
                        ]
                    )
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_wall_side_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_wall_side_column_2_stringvar.get()
                    ]
                    if not row_1[relevant_column_1] in allowed_values_1:
                        pipetally_2_paired.loc[
                            pipetally_2_paired.shape[0]
                        ] = pipetally_2_empty_row
                        continue
                    working_Data_2 = working_Data_2[
                        working_Data_2[relevant_column_2].isin(allowed_values_2)
                    ]
                if planned_step["coordinate_distance"][0]:
                    relevant_column_1_latitude = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_latitude_column_1_stringvar.get()
                    ]
                    relevant_column_1_longitude = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_longitude_column_1_stringvar.get()
                    ]
                    relevant_column_2_latitude = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_latitude_column_2_stringvar.get()
                    ]
                    relevant_column_2_longitude = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_longitude_column_2_stringvar.get()
                    ]
                    #  df['distance'] = 6367 * 2 * np.arcsin(np.sqrt(np.sin((np.radians(df['LAT']) - math.radians(37.2175900))/2)**2 + math.cos(math.radians(37.2175900)) * np.cos(np.radians(df['LAT'])) * np.sin((np.radians(df['LON']) - math.radians(-56.7213600))/2)**2))

                    working_Data_2["calculated_distance"] = (
                        6366155
                        * 2
                        * np.arcsin(
                            np.sqrt(
                                np.sin(
                                    (
                                        np.radians(
                                            working_Data_2[relevant_column_2_latitude]
                                        )
                                        - np.radians(row_1[relevant_column_1_latitude])
                                    )
                                    / 2
                                )
                                ** 2
                                + np.cos(np.radians(row_1[relevant_column_1_latitude]))
                                * np.cos(
                                    np.radians(
                                        working_Data_2[relevant_column_2_latitude]
                                    )
                                )
                                * np.sin(
                                    (
                                        np.radians(
                                            working_Data_2[relevant_column_2_longitude]
                                        )
                                        - np.radians(row_1[relevant_column_1_longitude])
                                    )
                                    / 2
                                )
                                ** 2
                            )
                        )
                    )
                    working_Data_2 = working_Data_2[
                        working_Data_2["calculated_distance"]
                        <= float(planned_step["coordinate_distance"][1])
                    ]
                    # working_Data_2 = working_Data_2[
                    #    working_Data_2.apply(lambda row_2, relevant_column_2_latitude=relevant_column_2_latitude, relevant_column_2_longitude=relevant_column_2_longitude, relevant_column_1_latitude=relevant_column_1_latitude, relevant_column_1_longitude=relevant_column_1_longitude, row_1=row_1: great_circle((row_2[relevant_column_2_latitude],row_2[relevant_column_2_longitude]),(row_1[relevant_column_1_latitude],row_1[relevant_column_1_longitude])).m, axis=1)
                    #    <= float(planned_step["coordinate_distance"][1])
                    # ]
                if planned_step["log_distance"][0]:
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_log_dist_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_log_dist_column_2_stringvar.get()
                    ]
                    working_Data_2 = working_Data_2[
                        (
                            working_Data_2[relevant_column_2] - row_1[relevant_column_1]
                        ).abs()
                        <= float(planned_step["log_distance"][1])
                    ]
                if planned_step["wt"][0]:
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_wt_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_wt_column_2_stringvar.get()
                    ]
                    working_Data_2 = working_Data_2[
                        (
                            working_Data_2[relevant_column_2] - row_1[relevant_column_1]
                        ).abs()
                        <= float(planned_step["wt"][1])
                    ]
                if planned_step["depth"][0]:
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_depth_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_depth_column_2_stringvar.get()
                    ]
                    working_Data_2 = working_Data_2[
                        (
                            working_Data_2[relevant_column_2] - row_1[relevant_column_1]
                        ).abs()
                        <= float(planned_step["depth"][1])
                    ]
                if planned_step["length"][0]:
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_length_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_length_column_2_stringvar.get()
                    ]
                    working_Data_2 = working_Data_2[
                        (
                            working_Data_2[relevant_column_2] - row_1[relevant_column_1]
                        ).abs()
                        <= float(planned_step["length"][1])
                    ]
                if planned_step["width"][0]:
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_width_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_width_column_2_stringvar.get()
                    ]
                    working_Data_2 = working_Data_2[
                        (
                            working_Data_2[relevant_column_2] - row_1[relevant_column_1]
                        ).abs()
                        <= float(planned_step["width"][1])
                    ]
                if planned_step["orientation"][0]:
                    if not "orientation_in_seconds" in self.pipetally_1_data.columns:
                        self.pipetally_1_data[
                            "orientation_in_seconds"
                        ] = self.pipetally_1_data[
                            self.column_letter_to_name_dictionary_1[
                                self.optionmenu_orientation_column_1_stringvar.get()
                            ]
                        ].apply(
                            self.convert_time_to_seconds
                        )
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_orientation_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_orientation_column_2_stringvar.get()
                    ]

                    working_Data_2 = working_Data_2[
                        (
                            working_Data_2["orientation_in_seconds"]
                            - row_1["orientation_in_seconds"]
                        ).abs()
                        <= self.convert_time_to_seconds(planned_step["orientation"][1])
                    ]
                if working_Data_2.shape[0] == 0:
                    pipetally_2_paired.loc[
                        pipetally_2_paired.shape[0]
                    ] = pipetally_2_empty_row
                elif working_Data_2.shape[0] == 1:
                    # print(row_1.transpose())
                    # pipetally_1_paired = pd.concat([pipetally_1_paired,row_1],ignore_index=True)
                    # pipetally_1_paired.loc[pipetally_1_paired.shape[0]] = row_1

                    pipetally_2_paired.loc[
                        pipetally_2_paired.shape[0]
                    ] = working_Data_2.iloc[0]
                    # pipetally_2_paired = pd.concat([pipetally_2_paired,working_Data_2],axis=0)
                # print(working_Data_2.head)
                # pd.concat([working_Data_2.head(1),row_1.transpose()],axis=1).to_excel('testa.xlsx',index=False)
                elif working_Data_2.shape[0] > 1:
                    relevant_column_1 = self.column_letter_to_name_dictionary_1[
                        self.optionmenu_log_dist_column_1_stringvar.get()
                    ]
                    relevant_column_2 = self.column_letter_to_name_dictionary_2[
                        self.optionmenu_log_dist_column_2_stringvar.get()
                    ]
                    working_Data_2 = working_Data_2.sort_values(
                        by=relevant_column_2,
                        key=lambda ord: (ord - row_1[relevant_column_1]).abs(),
                    )
                    # print(working_Data_2.head())

                    # pipetally_1_paired.loc[pipetally_1_paired.shape[0]] = row_1
                    # print(working_Data_2.head())
                    # print(working_Data_2.iloc[0])
                    # pdb.set_trace()
                    shrinking_Data_2.drop(
                        index=working_Data_2.iloc[0].name, inplace=True
                    )
                    pipetally_2_paired.loc[
                        pipetally_2_paired.shape[0]
                    ] = working_Data_2.iloc[0]
        # pd.concat([pipetally_1_paired,pipetally_1_paired],axis=1).to_excel("teasdat.xlsx")
        if "orientation_in_seconds" in self.pipetally_1_data.columns:
            self.pipetally_1_data.drop(columns="orientation_in_seconds")
        if "orientation_in_seconds" in pipetally_2_paired.columns:
            pipetally_2_paired.drop(columns="orientation_in_seconds")
        if "calculated_distance" in pipetally_2_paired.columns:
            pipetally_2_paired.drop(columns="calculated_distance")
        self.pipetally_1_data[" "] = np.nan  # add an empty column
        pd.concat([self.pipetally_1_data, pipetally_2_paired], axis=1).to_excel(
            self.output_path
        )
        # print(pipetally_2_paired.head())

    def convert_time_to_seconds(self, time):
        pattern = "([0-9]{1,2}):([0-9]{1,2}):?([0-9]{0,2})"
        found_match = re.search(pattern, str(time))
        if found_match:
            h, m, s = found_match.groups()
            if s:
                return int(s) + 60 * int(m) + 3600 * int(h)
            else:
                return 60 * int(m) + 3600 * int(h)
        else:
            return np.NaN


if __name__ == "__main__":
    root = tk.Tk()
    In_Line_Inspection_Result_Coupler(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
