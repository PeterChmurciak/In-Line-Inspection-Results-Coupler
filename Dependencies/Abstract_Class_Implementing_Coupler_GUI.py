# -*- coding: utf-8 -*-
"""
Abstract class facilitating the GUI for the In-Line Inspection Result Coupler app.
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
from abc import ABC, abstractmethod  # Creating abstract class
import tkinter as tk  # GUI creation tools
import os  # Working with paths
import string  # To provide alphabet letters


class In_Line_Inspection_Result_Coupler_GUI(tk.Frame, ABC):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master

        # Getting the path to app icon
        main_menu_icon_name = "pair_icon.ico"
        path_to_current_folder = os.path.dirname(os.path.abspath(__file__))
        self.path_to_main_menu_icon = os.path.join(
            path_to_current_folder, main_menu_icon_name
        )

        # Changing the icon
        self.master.iconbitmap(self.path_to_main_menu_icon)
        self.master.title("Inspection Result Coupler")
        self.master.resizable(False, False)

        self.frame_file_choice = tk.LabelFrame(self)

        self.button_pipetally_1 = tk.Button(self.frame_file_choice)
        self.button_pipetally_1.configure(
            font="{Segoe UI} 12 {bold}",
            text="PipeTally 1",
            command=self.button_pipetally_1_pressed,
        )
        self.button_pipetally_1.grid(column="0", padx="5", row="1")

        self.button_pipetally_2 = tk.Button(self.frame_file_choice)
        self.button_pipetally_2.configure(
            font="{Segoe UI} 12 {bold}",
            text="PipeTally 2",
            command=self.button_pipetally_2_pressed,
        )
        self.button_pipetally_2.grid(column="0", padx="5", row="2")

        self.button_output = tk.Button(self.frame_file_choice)
        self.button_output.configure(
            font="{Segoe UI} 12 {bold}",
            text="Výstup",
            command=self.button_output_pressed,
        )
        self.button_output.grid(column="0", padx="5", pady="0 5", row="3", sticky="ew")

        self.entry_pipetally_1_filename_stringvar = tk.StringVar(value="")
        self.entry_pipetally_1_filename = tk.Entry(self.frame_file_choice)
        self.entry_pipetally_1_filename.configure(
            justify="center",
            state="readonly",
            width="25",
            textvariable=self.entry_pipetally_1_filename_stringvar,
        )
        self.entry_pipetally_1_filename.grid(column="1", padx="5", row="1")
        self.pipetally_1_path = ""

        self.entry_pipetally_2_filename_stringvar = tk.StringVar(value="")
        self.entry_pipetally_2_filename = tk.Entry(self.frame_file_choice)
        self.entry_pipetally_2_filename.configure(
            justify="center",
            state="readonly",
            width="25",
            textvariable=self.entry_pipetally_2_filename_stringvar,
        )
        self.entry_pipetally_2_filename.grid(column="1", padx="5", row="2")
        self.pipetally_2_path = ""

        self.entry_output_filename_stringvar = tk.StringVar(value="")
        self.entry_output_filename = tk.Entry(self.frame_file_choice)
        self.entry_output_filename.configure(
            justify="center",
            state="readonly",
            width="25",
            textvariable=self.entry_output_filename_stringvar,
        )
        self.entry_output_filename.grid(column="1", padx="5", pady="0 5", row="3")
        self.output_path = ""

        self.entry_pipetally_1_range = tk.Entry(self.frame_file_choice)
        validate_command = (
            self.entry_pipetally_1_range.register(self.entry_range_validation),
            "%P",
        )
        self.entry_pipetally_1_range.configure(
            justify="center",
            validate="key",
            width="15",
            validatecommand=validate_command,
        )
        self.entry_pipetally_1_range.grid(column="2", padx="5", row="1")

        self.entry_pipetally_2_range = tk.Entry(self.frame_file_choice)
        validate_command = (
            self.entry_pipetally_2_range.register(self.entry_range_validation),
            "%P",
        )
        self.entry_pipetally_2_range.configure(
            justify="center",
            validate="key",
            width="15",
            validatecommand=validate_command,
        )
        self.entry_pipetally_2_range.grid(column="2", padx="5", row="2")

        self.ALPHABET = string.ascii_uppercase

        self.optionmenu_log_dist_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_log_dist_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_log_dist_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_log_dist_column_1.grid(
            column="3", padx="5", row="1", sticky="ew"
        )

        self.optionmenu_log_dist_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_log_dist_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_log_dist_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_log_dist_column_2.grid(
            column="3", padx="5", row="2", sticky="ew"
        )

        self.optionmenu_description_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_description_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_description_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_description_column_1.grid(
            column="4", padx="5", row="1", sticky="ew"
        )

        self.optionmenu_description_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_description_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_description_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_description_column_2.grid(
            column="4", padx="5", row="2", sticky="ew"
        )

        self.optionmenu_wt_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_wt_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_wt_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_wt_column_1.grid(column="5", padx="5", row="1", sticky="ew")

        self.optionmenu_wt_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_wt_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_wt_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_wt_column_2.grid(column="5", padx="5", row="2", sticky="ew")

        self.optionmenu_depth_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_depth_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_depth_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_depth_column_1.grid(column="6", padx="5", row="1", sticky="ew")

        self.optionmenu_depth_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_depth_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_depth_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_depth_column_2.grid(column="6", padx="5", row="2", sticky="ew")

        self.optionmenu_length_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_length_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_length_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_length_column_1.grid(column="7", padx="5", row="1", sticky="ew")

        self.optionmenu_length_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_length_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_length_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_length_column_2.grid(column="7", padx="5", row="2", sticky="ew")

        self.optionmenu_width_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_width_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_width_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_width_column_1.grid(column="8", padx="5", row="1", sticky="ew")

        self.optionmenu_width_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_width_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_width_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_width_column_2.grid(column="8", padx="5", row="2", sticky="ew")

        self.optionmenu_orientation_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_orientation_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_orientation_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_orientation_column_1.grid(
            column="9", padx="5", row="1", sticky="ew"
        )

        self.optionmenu_orientation_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_orientation_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_orientation_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_orientation_column_2.grid(
            column="9", padx="5", row="2", sticky="ew"
        )

        self.optionmenu_wall_side_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_wall_side_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_wall_side_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_wall_side_column_1.grid(
            column="10", padx="5", row="1", sticky="ew"
        )

        self.optionmenu_wall_side_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_wall_side_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_wall_side_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_wall_side_column_2.grid(
            column="10", padx="5", row="2", sticky="ew"
        )

        self.optionmenu_latitude_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_latitude_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_latitude_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_latitude_column_1.grid(
            column="11", padx="5", row="1", sticky="ew"
        )

        self.optionmenu_latitude_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_latitude_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_latitude_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_latitude_column_2.grid(
            column="11", padx="5", row="2", sticky="ew"
        )

        self.optionmenu_longitude_column_1_stringvar = tk.StringVar(value="")
        self.optionmenu_longitude_column_1 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_longitude_column_1_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_longitude_column_1.grid(
            column="12", padx="5", row="1", sticky="ew"
        )

        self.optionmenu_longitude_column_2_stringvar = tk.StringVar(value="")
        self.optionmenu_longitude_column_2 = tk.OptionMenu(
            self.frame_file_choice,
            self.optionmenu_longitude_column_2_stringvar,
            "",
            *self.ALPHABET
        )
        self.optionmenu_longitude_column_2.grid(
            column="12", padx="5", row="2", sticky="ew"
        )

        self.optionmenus_pipetally_1 = [
            self.optionmenu_log_dist_column_1,
            self.optionmenu_description_column_1,
            self.optionmenu_wt_column_1,
            self.optionmenu_depth_column_1,
            self.optionmenu_length_column_1,
            self.optionmenu_width_column_1,
            self.optionmenu_orientation_column_1,
            self.optionmenu_wall_side_column_1,
            self.optionmenu_latitude_column_1,
            self.optionmenu_longitude_column_1,
        ]

        self.optionmenu_stringvars_pipetally_1 = [
            self.optionmenu_log_dist_column_1_stringvar,
            self.optionmenu_description_column_1_stringvar,
            self.optionmenu_wt_column_1_stringvar,
            self.optionmenu_depth_column_1_stringvar,
            self.optionmenu_length_column_1_stringvar,
            self.optionmenu_width_column_1_stringvar,
            self.optionmenu_orientation_column_1_stringvar,
            self.optionmenu_wall_side_column_1_stringvar,
            self.optionmenu_latitude_column_1_stringvar,
            self.optionmenu_longitude_column_1_stringvar,
        ]

        self.optionmenus_pipetally_2 = [
            self.optionmenu_log_dist_column_2,
            self.optionmenu_description_column_2,
            self.optionmenu_wt_column_2,
            self.optionmenu_depth_column_2,
            self.optionmenu_length_column_2,
            self.optionmenu_width_column_2,
            self.optionmenu_orientation_column_2,
            self.optionmenu_wall_side_column_2,
            self.optionmenu_latitude_column_2,
            self.optionmenu_longitude_column_2,
        ]

        self.optionmenu_stringvars_pipetally_2 = [
            self.optionmenu_log_dist_column_2_stringvar,
            self.optionmenu_description_column_2_stringvar,
            self.optionmenu_wt_column_2_stringvar,
            self.optionmenu_depth_column_2_stringvar,
            self.optionmenu_length_column_2_stringvar,
            self.optionmenu_width_column_2_stringvar,
            self.optionmenu_orientation_column_2_stringvar,
            self.optionmenu_wall_side_column_2_stringvar,
            self.optionmenu_latitude_column_2_stringvar,
            self.optionmenu_longitude_column_2_stringvar,
        ]

        self.button_confirm = tk.Button(self.frame_file_choice)
        self.button_confirm.configure(
            font="{Segoe UI} 12 {bold}",
            text="Potvrdiť",
            command=self.button_confirm_pressed,
        )
        self.button_confirm.grid(
            column="2", columnspan="11", padx="5", pady="0 5", row="3", sticky="ew"
        )

        self.label_1 = tk.Label(self.frame_file_choice)
        self.label_1.configure(text="Vybrať súbor")
        self.label_1.grid(column="0", row="0")

        self.label_2 = tk.Label(self.frame_file_choice)
        self.label_2.configure(text="Názov súboru")
        self.label_2.grid(column="1", row="0")

        self.label_3 = tk.Label(self.frame_file_choice)
        self.label_3.configure(text="Rozsah tabuľky")
        self.label_3.grid(column="2", row="0")

        self.label_4 = tk.Label(self.frame_file_choice)
        self.label_4.configure(text="Log Dist.", width="8")
        self.label_4.grid(column="3", row="0")

        self.label_5 = tk.Label(self.frame_file_choice)
        self.label_5.configure(text="Description", width="8")
        self.label_5.grid(column="4", row="0")

        self.label_6 = tk.Label(self.frame_file_choice)
        self.label_6.configure(text="WT", width="8")
        self.label_6.grid(column="5", row="0")

        self.label_7 = tk.Label(self.frame_file_choice)
        self.label_7.configure(text="Depth", width="8")
        self.label_7.grid(column="6", row="0")

        self.label_8 = tk.Label(self.frame_file_choice)
        self.label_8.configure(text="Length", width="8")
        self.label_8.grid(column="7", row="0")

        self.label_9 = tk.Label(self.frame_file_choice)
        self.label_9.configure(text="Width", width="8")
        self.label_9.grid(column="8", row="0")

        self.label_10 = tk.Label(self.frame_file_choice)
        self.label_10.configure(text="Orientation", width="8")
        self.label_10.grid(column="9", row="0")

        self.label_11 = tk.Label(self.frame_file_choice)
        self.label_11.configure(text="Wall Side", width="8")
        self.label_11.grid(column="10", row="0")

        self.label_12 = tk.Label(self.frame_file_choice)
        self.label_12.configure(text="Latitude", width="8")
        self.label_12.grid(column="11", row="0")

        self.label_13 = tk.Label(self.frame_file_choice)
        self.label_13.configure(text="Longitude", width="8")
        self.label_13.grid(column="12", row="0")

        self.frame_file_choice.configure(
            font="{Segoe UI} 12 {bold}", text="Voľba súborov a špecifikovanie stĺpcov"
        )
        self.frame_file_choice.pack(
            expand="true", fill="both", padx="5", pady="0 5", side="top"
        )

        self.frame_setup_and_run = tk.LabelFrame(self)

        self.frame_numeric_criterion = tk.LabelFrame(self.frame_setup_and_run)

        self.checkbutton_log_dist_intvar = tk.IntVar(value=0)
        self.checkbutton_log_dist = tk.Checkbutton(self.frame_numeric_criterion)
        self.checkbutton_log_dist.configure(
            text="Log Distance",
            variable=self.checkbutton_log_dist_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_log_dist_tolerance, self.checkbutton_log_dist_intvar, "5"
            ),
        )
        self.checkbutton_log_dist.grid(column="0", padx="5", row="1", sticky="w")

        self.checkbutton_wt_intvar = tk.IntVar(value=0)
        self.checkbutton_wt = tk.Checkbutton(self.frame_numeric_criterion)
        self.checkbutton_wt.configure(
            text="WT",
            variable=self.checkbutton_wt_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_wt_tolerance, self.checkbutton_wt_intvar, "0"
            ),
        )
        self.checkbutton_wt.grid(column="0", padx="5", pady="5", row="2", sticky="w")

        self.checkbutton_depth_intvar = tk.IntVar(value=0)
        self.checkbutton_depth = tk.Checkbutton(self.frame_numeric_criterion)
        self.checkbutton_depth.configure(
            text="Depth",
            variable=self.checkbutton_depth_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_depth_tolerance, self.checkbutton_depth_intvar, "15"
            ),
        )
        self.checkbutton_depth.grid(column="0", padx="5", row="3", sticky="w")

        self.checkbutton_length_intvar = tk.IntVar(value=0)
        self.checkbutton_length = tk.Checkbutton(self.frame_numeric_criterion)
        self.checkbutton_length.configure(
            text="Length",
            variable=self.checkbutton_length_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_length_tolerance, self.checkbutton_length_intvar, "25"
            ),
        )
        self.checkbutton_length.grid(
            column="0", padx="5", pady="5", row="4", sticky="w"
        )

        self.checkbutton_width_intvar = tk.IntVar(value=0)
        self.checkbutton_width = tk.Checkbutton(self.frame_numeric_criterion)
        self.checkbutton_width.configure(
            text="Width",
            variable=self.checkbutton_width_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_width_tolerance, self.checkbutton_width_intvar, "25"
            ),
        )
        self.checkbutton_width.grid(column="0", padx="5", row="5", sticky="w")

        self.checkbutton_orientation_intvar = tk.IntVar(value=0)
        self.checkbutton_orientation = tk.Checkbutton(self.frame_numeric_criterion)
        self.checkbutton_orientation.configure(
            text="Orientation",
            variable=self.checkbutton_orientation_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_orientation_tolerance,
                self.checkbutton_orientation_intvar,
                "01:00",
            ),
        )
        self.checkbutton_orientation.grid(
            column="0", padx="5", pady="5", row="6", sticky="w"
        )

        self.checkbutton_coordinate_distance_intvar = tk.IntVar(value=0)
        self.checkbutton_coordinate_distance = tk.Checkbutton(
            self.frame_numeric_criterion
        )
        self.checkbutton_coordinate_distance.configure(
            text="Coordinate Distance",
            variable=self.checkbutton_coordinate_distance_intvar,
            state="disabled",
            command=lambda: self.change_entry_state(
                self.entry_coordinate_distance_tolerance,
                self.checkbutton_coordinate_distance_intvar,
                "5",
            ),
        )
        self.checkbutton_coordinate_distance.grid(
            column="0", padx="5", pady="0 5", row="7", sticky="w"
        )

        self.entry_log_dist_tolerance = tk.Entry(self.frame_numeric_criterion)
        validate_command = (
            self.entry_log_dist_tolerance.register(self.entry_number_validation),
            "%P",
        )
        self.entry_log_dist_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_log_dist_tolerance.grid(column="1", padx="5", row="1")

        self.entry_wt_tolerance = tk.Entry(self.frame_numeric_criterion)
        validate_command = (
            self.entry_wt_tolerance.register(self.entry_number_validation),
            "%P",
        )
        self.entry_wt_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_wt_tolerance.grid(column="1", padx="5", row="2")

        self.entry_depth_tolerance = tk.Entry(self.frame_numeric_criterion)
        validate_command = (
            self.entry_depth_tolerance.register(self.entry_number_validation),
            "%P",
        )
        self.entry_depth_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_depth_tolerance.grid(column="1", padx="5", row="3")

        self.entry_length_tolerance = tk.Entry(self.frame_numeric_criterion)
        validate_command = (
            self.entry_length_tolerance.register(self.entry_number_validation),
            "%P",
        )
        self.entry_length_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_length_tolerance.grid(column="1", padx="5", row="4")

        self.entry_width_tolerance = tk.Entry(self.frame_numeric_criterion)
        validate_command = (
            self.entry_width_tolerance.register(self.entry_number_validation),
            "%P",
        )
        self.entry_width_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_width_tolerance.grid(column="1", padx="5", row="5")

        self.entry_orientation_tolerance = tk.Entry(self.frame_numeric_criterion)
        validate_command = (
            self.entry_orientation_tolerance.register(
                self.entry_orientation_validation
            ),
            "%P",
        )
        self.entry_orientation_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_orientation_tolerance.grid(column="1", padx="5", row="6")

        self.entry_coordinate_distance_tolerance = tk.Entry(
            self.frame_numeric_criterion
        )
        validate_command = (
            self.entry_coordinate_distance_tolerance.register(
                self.entry_number_validation
            ),
            "%P",
        )
        self.entry_coordinate_distance_tolerance.configure(
            justify="center",
            validate="key",
            width="8",
            validatecommand=validate_command,
            state="disabled",
        )
        self.entry_coordinate_distance_tolerance.grid(
            column="1", padx="5", pady="0 5", row="7"
        )

        self.label_14 = tk.Label(self.frame_numeric_criterion)
        self.label_14.configure(text="Kritérium")
        self.label_14.grid(column="0", row="0")

        self.label_15 = tk.Label(self.frame_numeric_criterion)
        self.label_15.configure(text="Tolerancia")
        self.label_15.grid(column="1", row="0")

        self.label_16 = tk.Label(self.frame_numeric_criterion)
        self.label_16.configure(text="[m]")
        self.label_16.grid(column="3", padx="0 5", row="1", sticky="w")

        self.label_17 = tk.Label(self.frame_numeric_criterion)
        self.label_17.configure(text="[mm]")
        self.label_17.grid(column="3", padx="0 5", row="2", sticky="w")

        self.label_18 = tk.Label(self.frame_numeric_criterion)
        self.label_18.configure(text="[%WT]")
        self.label_18.grid(column="3", padx="0 5", row="3", sticky="w")

        self.label_19 = tk.Label(self.frame_numeric_criterion)
        self.label_19.configure(text="[mm]")
        self.label_19.grid(column="3", padx="0 5", row="4", sticky="w")

        self.label_20 = tk.Label(self.frame_numeric_criterion)
        self.label_20.configure(text="[mm]")
        self.label_20.grid(column="3", padx="0 5", row="5", sticky="w")

        self.label_21 = tk.Label(self.frame_numeric_criterion)
        self.label_21.configure(text="[hh:mm]")
        self.label_21.grid(column="3", padx="0 5", row="6", sticky="w")

        self.label_22 = tk.Label(self.frame_numeric_criterion)
        self.label_22.configure(text="[m]")
        self.label_22.grid(column="3", padx="0 5", pady="0 5", row="7", sticky="w")

        self.frame_numeric_criterion.configure(
            font="{Segoe UI} 11 {bold}", text="Numerické kriériá"
        )
        self.frame_numeric_criterion.grid(
            column="0", padx="5", pady="0 5", row="0", rowspan="2", sticky="nsew"
        )

        self.frame_categoric_criterion = tk.LabelFrame(self.frame_setup_and_run)

        self.checkbutton_description_intvar = tk.IntVar(value=0)
        self.checkbutton_description = tk.Checkbutton(self.frame_categoric_criterion)
        self.checkbutton_description.configure(
            text="Description",
            variable=self.checkbutton_description_intvar,
            state="disabled",
            command=lambda: self.change_listbox_state(
                self.listbox_description_1,
                self.listbox_description_2,
                self.checkbutton_description_intvar,
                self.listbox_description_1_data,
                (),
                self.listbox_description_2_data,
                (),
            ),
        )
        self.checkbutton_description.grid(
            column="0", padx="5", pady="5", row="1", sticky="nw"
        )

        self.checkbutton_wall_side_intvar = tk.IntVar(value=0)
        self.checkbutton_wall_side = tk.Checkbutton(self.frame_categoric_criterion)
        self.checkbutton_wall_side.configure(
            text="Wall Side",
            variable=self.checkbutton_wall_side_intvar,
            state="disabled",
            command=lambda: self.change_listbox_state(
                self.listbox_wall_side_1,
                self.listbox_wall_side_2,
                self.checkbutton_wall_side_intvar,
                self.listbox_wall_side_1_data,
                (),
                self.listbox_wall_side_2_data,
                (),
            ),
        )
        self.checkbutton_wall_side.grid(column="0", padx="5", row="2", sticky="nw")

        self.label_23 = tk.Label(self.frame_categoric_criterion)
        self.label_23.configure(text="Kritérium")
        self.label_23.grid(column="0", pady="0 4", row="0", sticky="n")

        self.label_24 = tk.Label(self.frame_categoric_criterion)
        self.label_24.configure(text="Dovolené kategórie\nPipeTally 1")
        self.label_24.grid(column="1", pady="0 4", row="0", sticky="n")

        self.label_25 = tk.Label(self.frame_categoric_criterion)
        self.label_25.configure(text="Dovolené kategórie\nPipeTally 2")
        self.label_25.grid(column="2", pady="0 4", row="0", sticky="n")

        self.listbox_description_1_frame = tk.Frame(self.frame_categoric_criterion)
        self.listbox_description_1 = tk.Listbox(self.listbox_description_1_frame)
        self.listbox_description_1.configure(
            activestyle="none",
            height="5",
            selectmode="multiple",
            width="24",
            state="disabled",
            exportselection=False,
        )
        self.listbox_description_1.pack(expand="true", fill="both", side="left")

        self.scrollbar_listbox_description_1 = tk.Scrollbar(
            self.listbox_description_1_frame
        )
        self.scrollbar_listbox_description_1.configure(
            orient="vertical", command=self.listbox_description_1.yview
        )
        self.listbox_description_1.configure(
            yscrollcommand=self.scrollbar_listbox_description_1.set
        )
        self.scrollbar_listbox_description_1.pack(
            expand="true", fill="both", side="right"
        )
        self.listbox_description_1_frame.grid(column="1", padx="5", pady="5", row="1")

        self.listbox_description_2_frame = tk.Frame(self.frame_categoric_criterion)
        self.listbox_description_2 = tk.Listbox(self.listbox_description_2_frame)
        self.listbox_description_2.configure(
            activestyle="none",
            height="5",
            selectmode="multiple",
            width="24",
            state="disabled",
            exportselection=False,
        )
        self.listbox_description_2.pack(expand="true", fill="both", side="left")

        self.scrollbar_listbox_description_2 = tk.Scrollbar(
            self.listbox_description_2_frame
        )
        self.scrollbar_listbox_description_2.configure(
            orient="vertical", command=self.listbox_description_2.yview
        )
        self.listbox_description_2.configure(
            yscrollcommand=self.scrollbar_listbox_description_2.set
        )
        self.scrollbar_listbox_description_2.pack(
            expand="true", fill="both", side="right"
        )
        self.listbox_description_2_frame.grid(column="2", padx="5", pady="5", row="1")

        self.listbox_wall_side_1_frame = tk.Frame(self.frame_categoric_criterion)
        self.listbox_wall_side_1 = tk.Listbox(self.listbox_wall_side_1_frame)
        self.listbox_wall_side_1.configure(
            activestyle="none",
            height="5",
            selectmode="multiple",
            width="24",
            state="disabled",
            exportselection=False,
        )
        self.listbox_wall_side_1.pack(expand="true", fill="both", side="left")

        self.scrollbar_listbox_wall_side_1 = tk.Scrollbar(
            self.listbox_wall_side_1_frame
        )
        self.scrollbar_listbox_wall_side_1.configure(
            orient="vertical", command=self.listbox_wall_side_1.yview
        )
        self.listbox_wall_side_1.configure(
            yscrollcommand=self.scrollbar_listbox_wall_side_1.set
        )
        self.scrollbar_listbox_wall_side_1.pack(
            expand="true", fill="both", side="right"
        )
        self.listbox_wall_side_1_frame.grid(column="1", padx="5", pady="5", row="2")

        self.listbox_wall_side_2_frame = tk.Frame(self.frame_categoric_criterion)
        self.listbox_wall_side_2 = tk.Listbox(self.listbox_wall_side_2_frame)
        self.listbox_wall_side_2.configure(
            activestyle="none",
            height="5",
            selectmode="multiple",
            width="24",
            state="disabled",
            exportselection=False,
        )
        self.listbox_wall_side_2.pack(expand="true", fill="both", side="left")

        self.scrollbar_listbox_wall_side_2 = tk.Scrollbar(
            self.listbox_wall_side_2_frame
        )
        self.scrollbar_listbox_wall_side_2.configure(
            orient="vertical", command=self.listbox_wall_side_2.yview
        )
        self.listbox_wall_side_2.configure(
            yscrollcommand=self.scrollbar_listbox_wall_side_2.set
        )
        self.scrollbar_listbox_wall_side_2.pack(
            expand="true", fill="both", side="right"
        )
        self.listbox_wall_side_2_frame.grid(column="2", padx="5", pady="5", row="2")

        self.frame_categoric_criterion.configure(
            font="{Segoe UI} 11 {bold}", text="Kategorické kriériá"
        )
        self.frame_categoric_criterion.grid(
            column="1", pady="0 5", row="0", rowspan="2", sticky="nsew"
        )

        self.frame_multistep_pairing = tk.LabelFrame(self.frame_setup_and_run)

        self.listbox_list_of_steps_frame = tk.Frame(self.frame_multistep_pairing)
        self.listbox_list_of_steps = tk.Listbox(self.listbox_list_of_steps_frame)
        self.listbox_list_of_steps.configure(
            activestyle="none",
            height="9",
            selectmode="single",
            width="16",
            exportselection=False,
        )
        self.listbox_list_of_steps.insert(0, "Krok číslo 1")
        self.listbox_list_of_steps.selection_set(0)
        self.listbox_list_of_steps.pack(expand="true", fill="both", side="left")
        self.listbox_list_of_steps.bind(
            "<<ListboxSelect>>", self.button_load_step_pressed
        )

        self.scrollbar_listbox_list_of_steps = tk.Scrollbar(
            self.listbox_list_of_steps_frame
        )
        self.scrollbar_listbox_list_of_steps.configure(
            orient="vertical", command=self.listbox_list_of_steps.yview
        )
        self.listbox_list_of_steps.configure(
            yscrollcommand=self.scrollbar_listbox_list_of_steps.set
        )
        self.scrollbar_listbox_list_of_steps.pack(
            expand="true", fill="both", side="right"
        )
        self.listbox_list_of_steps_frame.grid(
            column="0", padx="5 0", pady="5 0", row="1", rowspan="4"
        )

        self.button_add_step = tk.Button(self.frame_multistep_pairing)
        self.button_add_step.configure(
            font="{Segoe UI} 12 {bold}",
            text="Pridať krok",
            command=self.button_add_step_pressed,
        )
        self.button_add_step.grid(
            column="1", padx="6", pady="5 0", row="1", sticky="ew"
        )

        self.button_save_step = tk.Button(self.frame_multistep_pairing)
        self.button_save_step.configure(
            font="{Segoe UI} 12 {bold}",
            text="Uložiť krok",
            command=self.button_save_step_pressed,
        )
        self.button_save_step.grid(column="1", padx="6", row="2", sticky="ew")

        self.button_load_step = tk.Button(self.frame_multistep_pairing)
        self.button_load_step.configure(
            font="{Segoe UI} 12 {bold}",
            text="Načítať krok",
            command=self.button_load_step_pressed,
        )
        self.button_load_step.grid(column="1", padx="6", row="3", sticky="ew")

        self.button_delete_step = tk.Button(self.frame_multistep_pairing)
        self.button_delete_step.configure(
            font="{Segoe UI} 12 {bold}",
            text="Odobrať krok",
            command=self.button_delete_step_pressed,
        )
        self.button_delete_step.grid(column="1", padx="6", row="4", sticky="ew")

        self.label_26 = tk.Label(self.frame_multistep_pairing)
        self.label_26.configure(text="Zoznam krokov")
        self.label_26.grid(column="0", row="0")

        self.label_27 = tk.Label(self.frame_multistep_pairing)
        self.label_27.configure(text="Možnosti")
        self.label_27.grid(column="1", row="0")

        self.frame_multistep_pairing.configure(
            font="{Segoe UI} 11 {bold}",
            height="200",
            text="Viackrokové párovanie",
            width="200",
        )
        self.frame_multistep_pairing.grid(column="2", padx="5", row="0", sticky="nsew")

        self.button_start_pairing = tk.Button(self.frame_setup_and_run)
        self.button_start_pairing.configure(
            font="{Segoe UI} 12 {bold}",
            text="Spustiť párovanie",
            command=self.button_start_pairing_pressed,
        )
        self.button_start_pairing.grid(
            column="2", padx="5", pady="5 6", row="1", sticky="nsew"
        )

        self.frame_setup_and_run.configure(
            font="{Segoe UI} 12 {bold}",
            text="Nastavenie kritérií a spustenie párovania",
        )
        self.frame_setup_and_run.pack(
            expand="true", fill="both", padx="5", pady="0 5", side="top"
        )

        self.template_dictionary = {
            "log_distance": (0, ""),
            "wt": (0, ""),
            "depth": (0, ""),
            "length": (0, ""),
            "width": (0, ""),
            "orientation": (0, ""),
            "coordinate_distance": (0, ""),
            "description": (0, (), (), (), ()),
            "wall_side": (0, (), (), (), ()),
        }
        self.dictionary_of_steps = {"Krok číslo 1": self.template_dictionary.copy()}

        self.checkbuttons = [
            self.checkbutton_log_dist,
            self.checkbutton_description,
            self.checkbutton_wt,
            self.checkbutton_depth,
            self.checkbutton_length,
            self.checkbutton_width,
            self.checkbutton_orientation,
            self.checkbutton_wall_side,
            self.checkbutton_coordinate_distance,
        ]

    @abstractmethod
    def button_pipetally_1_pressed(self):
        pass

    @abstractmethod
    def button_pipetally_2_pressed(self):
        pass

    @abstractmethod
    def button_output_pressed(self):
        pass

    @abstractmethod
    def entry_range_validation(self, P):
        pass

    @abstractmethod
    def reset_setup_frame(self):
        pass

    @abstractmethod
    def button_confirm_pressed(self):
        pass

    @abstractmethod
    def entry_number_validation(self, P):
        pass

    @abstractmethod
    def entry_orientation_validation(self, P):
        pass

    @abstractmethod
    def button_add_step_pressed(self):
        pass

    @abstractmethod
    def button_save_step_pressed(self):
        pass

    @abstractmethod
    def button_load_step_pressed(self, event=None):
        pass

    @abstractmethod
    def button_delete_step_pressed(self):
        pass

    @abstractmethod
    def change_entry_state(self, entry, checkbutton_intvar, custom_value=""):
        pass

    @abstractmethod
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
        pass

    @abstractmethod
    def button_start_pairing_pressed(self):
        pass

    @abstractmethod
    def convert_time_to_seconds(self, time):
        pass
