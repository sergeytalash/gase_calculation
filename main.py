import datetime
import json
import os
import tkinter as tk
from datetime import timedelta

import matplotlib
import numpy as np
import pandas as pd

matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

pd.set_option("display.max_rows", None,
              "display.max_columns", None,
              "display.width", 140)


class Chart(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent, background="black")
        self.fig = Figure(figsize=(10, 7), dpi=100)
        self.a1 = self.fig.add_subplot(111)
        self.a2 = self.a1
        self.a3 = self.a1
        self.b1 = self.a1
        self.canvas = FigureCanvasTkAgg(self.fig, self)
        self.canvas.draw()
        self.canvas.get_tk_widget().grid()
        self.grid(column=1, row=0)

    def set_legend(self, legend):
        """
        Creates a legend on the canvas field
        Args:
            legend (list): List of legend items

        Returns:
            None
        """
        self.fig.legend(legend, loc="upper left")

    def set_plot(self, x, y, ax, line_width=0, marker='o'):
        """
        Creates data items on the canvas field
        Args:
            x (iter): Data index array
            y (iter): Data values array
            ax (str): Target ax to plot on
            line_width (int): Width of the line

        Returns:
            None
        """
        if ax == "2":
            self.a2 = self.a1.twinx()
        if ax == "3":
            self.a3 = self.a1.twinx()
        if ax == "11":
            self.b1 = self.a1.twinx().twiny()
        axes = {"1": self.a1, "2": self.a2, "3": self.a3, "11": self.b1}
        color = {"1": "black", "2": "blue", "3": "red", "11": "green"}
        if len(x) == len(y):
            current_ax = axes.get(ax, self.a1)
            current_ax.tick_params(labelcolor=color.get(ax, color["1"]))
            current_ax.plot(x, y,
                            color=color.get(ax, color["1"]),
                            marker=marker,
                            linewidth=line_width,
                            alpha=1,
                            )

    def clear(self):
        for item in self.canvas.get_tk_widget().find_all():
            self.canvas.get_tk_widget().delete(item)

    def show(self, data, mode, line_width=2, marker=None):
        if mode == "data.index, data.MPVPosition":
            self.set_plot(data.index, data.MPVPosition, ax="1", line_width=line_width, marker=marker)
            self.set_plot(data.index, data.CH4_dry, ax="2", line_width=line_width, marker=marker)
            self.set_plot(data.index, data.CO2_dry, ax="3", line_width=line_width, marker=marker)
            legend = ["MPVPosition", "CH4_dry", "CO2_dry"]
        elif mode == "data2.index, data2.MPVPosition.mean":
            data2 = pd.concat(data.values())
            # self.set_plot(data2.index, data2.MPVPosition["mean"], ax="1", line_width=line_width)
            self.set_plot(data2.index, data2.CH4_dry["mean"], ax="1", line_width=line_width, marker=marker)
            self.set_plot(data2.index, data2.CO2_dry["mean"], ax="2", line_width=line_width, marker=marker)
            legend = ["CH4_dry", "CO2_dry"]
        elif mode == "calibrated_data":
            axes = {'CH4+': "1", 'CO2+': "11"}
            legend = []
            for i, cycle in enumerate(data):
                for j, gas in enumerate(['CH4+', 'CO2+'], start=1):
                    self.set_plot(data[i][gas]['measured'],
                                  data[i][gas]['assigned'],
                                  ax=axes[gas],
                                  line_width=line_width)
                    legend.append(f"{gas}_{i}")
        else:
            if isinstance(mode, list):
                for i, column in enumerate(mode, start=1):
                    self.set_plot(data["date_time"],
                                  data[column],
                                  ax=str(i),
                                  line_width=line_width)
                legend = mode
            elif isinstance(mode, str):
                self.set_plot(data["date_time"],
                              data[mode],
                              ax="1",
                              line_width=line_width)
                legend = [mode]
            else:
                legend = ["None"]
        self.set_legend(legend=legend)
        self.update()


class Calculate:
    def __init__(self):
        self.baloon_concentration_file = "baloon_concentration.json"
        self.concentration = self.open_concentration()
        self.data_folder = "data"
        self.columns = ("DATETIME", "MPVPosition", "CO2_dry", "CH4_dry")
        self.CO2_std_limit = 0.02
        self.CH4_std_limit = 0.0002
        self.round_calculated_values = {"CH4+": 5, "CO2+": 2}
        self.one_minute_resample_filename = "table_by_1_minute.xlsx"
        self.ch4_table_filename = "table_ch4.xlsx"
        self.co2_table_filename = "table_co2.xlsx"

    def open_concentration(self):
        """
        Reads concentration json file
        Returns:
            dict: Concentration data
        """
        with open(self.baloon_concentration_file, encoding="utf8") as fr:
            return json.load(fr)

    @staticmethod
    def open_data_file(file_name):
        """
        Reads data file
        Args:
            file_name (str): Target data file

        Returns:
            pd.DataFrame: Data table
        """
        file = pd.read_csv(file_name, sep="\s+")
        return pd.DataFrame(file)

    def open_all_data_files(self):
        dfs = []
        for file in os.listdir(self.data_folder):
            dfs.append(self.open_data_file(os.path.join(self.data_folder, file)))
        return pd.concat(dfs)

    def get_data(self):
        """
        Reads all files in data folder and adds datetime object to the DATETIME column.
        Uses DATE and TIME columns as sources
        Returns:
            pd.DataFrame: Data table with DATETIME column
        """
        data = self.open_all_data_files()
        data["DATETIME"] = pd.to_datetime(data['DATE'] + ' ' + data['TIME'])
        data.set_index("DATETIME")
        data = data.sort_values(by="DATETIME")
        new_columns = {}
        for old_column in self.columns:
            for new_column in data.columns:
                if old_column in new_column:
                    new_columns[new_column] = old_column
        data = data.rename(columns=new_columns)
        return data

    @staticmethod
    def multiply_1000(data, column_name):
        """
        Multiplies ch4 measurement values on 1000
        Returns:
            pd.DataFrame: Data table with CH4_dry column
        """
        data[column_name] = data[column_name] * 1000
        return data

    def group_by_mvp_position(self, data):
        """
        Groups data by mvp position.
        (*) Removes data with MVPposition which is different from defined in concentration file
        Args:
            data (pd.DataFrame): Data table

        Returns:
            dict: {i: {"DATETIME": [], "MPVPosition": [], "CO2_dry": [], "CH4_dry": []},.. }
        """
        first_mvp = data["MPVPosition"].head(n=1).values[0]
        if "DATETIME" in data.columns:
            mvp_line_index = 1
            tmp_data = data.loc[:, self.columns]
        else:
            mvp_line_index = 0
            tmp_data = data
        new_data = {}
        index = 0
        for data_index, line in zip(tmp_data.index, tmp_data.values):  # "DATETIME", "MPVPosition", "CO2_dry", "CH4_dry"
            mvp = float(line[mvp_line_index])
            line[mvp_line_index] = mvp
            line = line.tolist()

            if mvp_line_index == 0:
                line.insert(0, data_index)
            # (*)
            if mvp not in [float(k) for k in self.concentration.keys()]:
                continue
            if line[1] != first_mvp:
                index += 1
                first_mvp = line[1]
            new_data.setdefault(index, []).append(line)
        return new_data

    def make_dataframe_dict(self, data, part=1):
        """
        Replaces dict of lists with pd.DataFrame
        Args:
            part (int): Part of calculations
            data (dict): Data with dict of lists

        Returns:
            dict: {i: pd.DataFrame,.. }
        """
        dataframe_dict = {}
        if part == 1:
            columns = self.columns
        elif part == 2:
            columns = pd.MultiIndex.from_tuples(
                [('DATETIME', 'mean'),
                 ('MPVPosition', 'mean'),
                 ('CH4_dry', 'mean'),
                 ('CH4_dry', 'std'),
                 ('CO2_dry', 'mean'),
                 ('CO2_dry', 'std')],
                )
        else:
            columns = []

        for index, tmp_data in data.items():
            df = pd.DataFrame(data=tmp_data, columns=columns)
            df.set_index(self.columns[0])
            dataframe_dict[index] = df

        return dataframe_dict

    def resample_by_1_minute(self, data):
        """
        Resamples data by minute
        Args:
            data (dict): Data with dict of pd.DataFrame

        Returns:
            dict: Data with dict of pd.DataFrame
        """
        new_data = {}
        for index, array in data.items():
            if array.size > 10:

                df = array.resample(timedelta(minutes=1), on='DATETIME').agg(
                    {'MPVPosition': 'mean',
                     'CH4_dry': ['mean', 'std'],
                     'CO2_dry': ['mean', 'std'],
                     })
                if (df.index[df.CH4_dry['std'] < self.CH4_std_limit].size > 2 and
                    df.index[df.CO2_dry['std'] < self.CO2_std_limit].size > 2):
                    new_data[index] = df
                    # print(f"Size: {len(df.index)} -> "
                    #       f"CH4_dry: {len(df.index[df.CH4_dry['std'] < self.CH4_std_limit])} "
                    #       f"CO2_dry: {len(df.index[df.CO2_dry['std'] < self.CO2_std_limit])}"
                    #       )
        return new_data

    @staticmethod
    def save_to_excel(data, filename):
        if isinstance(data, dict):
            data2 = pd.concat(data.values())
        else:
            data2 = data
        try:
            data2.to_excel(filename)
        except Exception as err:
            print(f"Error in saving to file {filename}:\n{err}\nProbably file is already in use.")
            filename = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S_") + filename
            data2.to_excel(filename)
        finally:
            print(f"File has been created {filename}.")

    @staticmethod
    def read_from_excel(filename="table_by_1_minute.xlsx"):
        """
        Asks user to update excel file and then reads it
        Args:
            filename (str): Part of the file name

        Returns:
            pd.DataFrame: or None
        """
        for file in os.listdir():
            if filename in file:
                try:
                    input(f"Вы можете редактировать промежуточный файл '{file}'\n"
                          f"Сохраните его и нажмите Enter чтобы продолжить...")
                    return pd.read_excel(file, header=[0, 1], index_col=0)
                except Exception as err:
                    print(err)

    @staticmethod
    def take_last(data):
        """
        Takes only a half of the table from the end
        Args:
            data (dict): Data with dict of pd.DataFrame

        Returns:
            dict: Data with dict of pd.DataFrame
        """
        new_data = {}
        for index, array in data.items():
            df = array[-8:]
            df.reset_index(inplace=True)
            new_data[index] = df
        return new_data

    @staticmethod
    def make_mean(data):
        """
        Takes only a half of the table from the end
        Args:
            data (dict): Data with dict of pd.DataFrame

        Returns:
            dict: Data with dict of pd.DataFrame
        """
        new_data = {}
        for index, array in data.items():
            tmp_array = array.iloc[:, 2:]
            new_data[index] = tmp_array.mean()
            new_data[index]["datetime"] = array["DATETIME"]["mean"][0]
        return new_data

    def calc_coefficients(self, data):
        """
        Calculates polynomial coefficients for calibration gases
        Args:
            data (pd.DataFrame):

        Returns:

        """
        data_dict = data.to_dict(orient='split')
        measure_cycles = {}
        cycle_i = 0
        cycle = 0
        for line, date_time in zip(data_dict["data"], data_dict["index"]):
            if cycle_i < len(self.concentration):
                cycle_i += 1
            else:
                cycle_i = 1
                cycle += 1
            measure_cycles.setdefault(cycle, {}).setdefault(
                "data", {})[str(float(line[0]))] = [line[1], line[3]]  # MVP,CH4,CO2
            measure_cycles.setdefault(cycle, {}).setdefault(
                "std", {})[str(float(line[0]))] = [line[2], line[4]]  # MVP,CH4_std,CO2_std
            measure_cycles.setdefault(cycle, {}).setdefault(
                "date_time", {})[str(float(line[0]))] = date_time
        measure_cycles = {i: cycle for i, cycle in measure_cycles.items() if
                          len(cycle["data"]) == len(self.concentration)}
        mvp_for_calibration = {k: v for k, v in self.concentration.items() if v["for_calibration"]}

        calibrated_gases = []
        for i, cycle in measure_cycles.items():
            gases_dict = {}
            for mvp, values in mvp_for_calibration.items():
                for j, gas in enumerate(["CH4+", "CO2+"]):
                    gases_dict.setdefault(
                        gas, {}).setdefault(
                        "measured", []).append(cycle["data"][mvp][j])
                    gases_dict.setdefault(
                        gas, {}).setdefault(
                        "std", []).append(cycle["std"][mvp][j])
                    gases_dict.setdefault(
                        gas, {}).setdefault(
                        "assigned", []).append(values[gas])
                    gases_dict.setdefault(
                        gas, {}).setdefault(
                        "date_time", []).append(cycle["date_time"][mvp])
            # concentration CO2(measured CO2)
            # concentration CH4(measured CH4)
            ch4_coeffs = np.polyfit(gases_dict["CH4+"]["measured"], gases_dict["CH4+"]["assigned"], deg=1)
            co2_coeffs = np.polyfit(gases_dict["CO2+"]["measured"], gases_dict["CO2+"]["assigned"], deg=1)
            gases_dict["CH4+"]["coefficients"] = list(ch4_coeffs)
            gases_dict["CO2+"]["coefficients"] = list(co2_coeffs)
            calibrated_gases.append(gases_dict)
        mvp_not_for_calibration = {k: v for k, v in self.concentration.items() if not v["for_calibration"]}

        gases = []
        for i, cycle in measure_cycles.items():
            gases_dict = {}
            for mvp, values in mvp_not_for_calibration.items():
                balloon_name = values["name"]
                for j, gas in enumerate(["CH4+", "CO2+"]):
                    measured_gas_value = cycle["data"][mvp][j]
                    gases_dict.setdefault(
                        balloon_name, {}).setdefault(
                        gas, {})["measured"] = measured_gas_value
                    gases_dict.setdefault(
                        balloon_name, {}).setdefault(
                        gas, {})["std"] = cycle["std"][mvp][j]
                    coefficients = calibrated_gases[i][gas]['coefficients']
                    calculated_value = coefficients[0] * measured_gas_value + coefficients[1]
                    gases_dict.setdefault(
                        balloon_name, {}).setdefault(
                        gas, {})["calculated"] = calculated_value
                    gases_dict.setdefault(
                        balloon_name, {}).setdefault(
                        gas, {})["date_time"] = cycle["date_time"][mvp]
            gases.append(gases_dict)

        return calibrated_gases, gases

    @staticmethod
    def self_check(cd):
        # coeffs = {"CH4+": cd[0]['CH4+']['coefficients'],
        #           "CO2+": cd[0]['CO2+']['coefficients']}
        for i, cycle in enumerate(cd):
            coeffs = {"CH4+": cd[i]['CH4+']['coefficients'],
                      "CO2+": cd[i]['CO2+']['coefficients']}
            for name, gas in cycle.items():
                calculated = {}
                for measure in gas['measured']:
                    calculated_value = coeffs[name][0] * measure + coeffs[name][1]
                    calculated.setdefault("calculated", []).append(calculated_value)
                gas.update(calculated)
        return cd

    def reformat_calibrated_gases(self, data):
        co2_list = []
        ch4_list = []
        calibration_gases_names = [v["name"] for v in self.concentration.values()
                                   if v["for_calibration"]]
        for cycle in data:
            for name, gases_n1_n2 in cycle.items():
                gases = [{}, {}]
                for item, array in gases_n1_n2.items():
                    for i, d in enumerate(array):
                        gases[i]["name"] = calibration_gases_names[i]
                        gases[i][item] = d

                if name == 'CH4+':
                    ch4_list.extend(gases)

                if name == 'CO2+':
                    co2_list.extend(gases)

        return co2_list, ch4_list

    @staticmethod
    def reformat_common_gases(data):
        co2_list = []
        ch4_list = []
        for cycle in data:
            for n, gas_pair in cycle.items():
                for name, gas in gas_pair.items():
                    gas["name"] = n
                    if name == 'CH4+':
                        ch4_list.append(gas)
                    if name == 'CO2+':
                        co2_list.append(gas)
        return co2_list, ch4_list

    @staticmethod
    def make_table(data):
        data = pd.DataFrame(data)
        data = data.reindex(['date_time', 'name', 'measured', "std", 'assigned', 'calculated', 'coefficients'], axis=1)
        data = data.sort_values(by="date_time")
        data = data.reset_index(drop=True)
        return data


class MainApp(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)

        main = Calculate()
        # TODO: Allow working with 3 calibration gases
        df = main.get_data()
        data_dict = main.group_by_mvp_position(df)
        data_dict = main.make_dataframe_dict(data_dict, part=1)
        data_dict = main.resample_by_1_minute(data_dict)

        main.save_to_excel(data_dict, filename=main.one_minute_resample_filename)
        Chart(self).show(data_dict, mode="data2.index, data2.MPVPosition.mean", line_width=1, marker=None)

        df = main.read_from_excel()

        data_dict = main.group_by_mvp_position(df)
        data_dict = main.make_dataframe_dict(data_dict, part=2)
        data_dict = main.take_last(data_dict)
        # print(data_dict)
        # Chart(self).show(data, mode="data2.index, data2.MPVPosition.mean")
        data_dict = main.make_mean(data_dict)
        df = pd.concat(data_dict.values())

        df = pd.DataFrame(dict(MPVPosition=df.MPVPosition["mean"].values,
                               CH4_dry=df.CH4_dry["mean"].values,
                               CH4_dry_std=df.CH4_dry["std"].values,
                               CO2_dry=df.CO2_dry["mean"].values,
                               CO2_dry_std=df.CO2_dry["std"].values),
                          index=df.datetime.values).sort_index()
        # print(df)
        # Chart(self).show(data, mode="data.index, data.MPVPosition")
        calibrated_data, gases = main.calc_coefficients(df)
        # Chart(self).show(calibrated_data, mode="calibrated_data")

        recalculated_calibrated_gases = main.self_check(calibrated_data)

        co2_1, ch4_1 = main.reformat_calibrated_gases(recalculated_calibrated_gases)
        co2_2, ch4_2 = main.reformat_common_gases(gases)
        co2 = co2_1 + co2_2
        ch4 = ch4_1 + ch4_2

        co2_table = main.make_table(co2)
        co2_table = co2_table
        main.save_to_excel(co2_table, filename=main.co2_table_filename)

        ch4_table = main.make_table(ch4)
        for column in ["measured", "std", "calculated", "assigned"]:
            ch4_table = main.multiply_1000(ch4_table, column_name=column)
        main.save_to_excel(ch4_table, filename=main.ch4_table_filename)

        print("\nCO2")
        print(co2_table)
        print("\nCH4")
        print(ch4_table)
        # Chart(self).show(co2_table, mode=["calculated", "measured", "name"])
        # Chart(self).show(ch4_table, mode="measured")


if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
    input("\nPress Enter to exit...")
