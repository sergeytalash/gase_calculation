import os
import sys
import tkinter as tk
import warnings
from datetime import datetime as dt
from datetime import timedelta
from pprint import pprint

import matplotlib
import numpy as np
import openpyxl


warnings.simplefilter(action='ignore', category=UserWarning)
import pandas as pd

matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

pd.set_option("display.max_rows", None,
              "display.max_columns", None,
              "display.width", 140)

data_folder = "data"
balloon_concentration_file = os.path.join(data_folder, "balloon_concentration.xlsx")
one_minute_resample_filename = os.path.join(data_folder, "_table_by_1_minute_no_std.xlsx")
ch4_table_filename = os.path.join(data_folder, "_table_ch4.xlsx")
co2_table_filename = os.path.join(data_folder, "_table_co2.xlsx")


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
            marker (str, optional): Marker type

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
                            alpha=1)

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
    def __init__(self, std, no_flask):
        self.filter_by_std = std
        self.no_flask = no_flask
        self.concentration = self.open_concentration_xlsx()
        self.columns = ("DATETIME", "MPVPosition", "CO2_dry", "CH4_dry")
        self.CO2_std_limit = 0.02
        self.CH4_std_limit = 0.0002

    def open_concentration_xlsx(self):
        """
        Reads concentration json file
        Returns:
            dict: Concentration data
        """
        concentration_dict = {}
        wb = openpyxl.load_workbook(balloon_concentration_file)
        ws = wb.active
        data = ws.iter_rows(values_only=True)
        next(data)
        for d in data:
            if all(d[:2]):
                concentration_dict[str(d[0])] = {
                    "name": str(d[1]),
                    "CO2+": float(d[2]),
                    "CH4+": float(d[3]),
                    "for_calibration": str(d[4])}
        return concentration_dict

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
        all_files = []
        for path, folders, files in os.walk(os.path.join(data_folder, "DataLog")):
            for file in files:
                if not file.endswith(".dat"):
                    continue
                full_path = os.path.join(path, file)
                if os.path.isfile(full_path):
                    all_files.append(full_path)
        all_files_len = len(all_files)
        dfs = []
        for i, file in enumerate(all_files):
            dfs.append(self.open_data_file(file))
            if i % 1000 == 0 or i + 1 == all_files_len:
                print(f"Loading files [{i} / {all_files_len}]")
        return pd.concat(dfs)

    def get_data(self):
        """
        Reads all files in data folder and adds datetime object to the DATETIME column.
        Uses DATE and TIME columns as sources
        Returns:
            pd.DataFrame: Data table with DATETIME column
        """
        cache_file = "data.cache"
        t = dt.now()
        if not os.path.exists(cache_file):
            data = self.open_all_data_files()
            with open(cache_file, "w") as fw:
                data.to_csv(fw)
        else:
            with open(cache_file, "r") as fr:
                data = pd.read_csv(fr)
        print(dt.now() - t)


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

    def group_by_mpv_position(self, data):
        """
        Groups data by mpv position.
        (*) Removes data with mpv position which is different from defined in concentration file
        Args:
            data (pd.DataFrame): Data table

        Returns:
            dict: {i: {"DATETIME": [], "MPVPosition": [], "CO2_dry": [], "CH4_dry": []},.. }
        """
        first_mpv = data["MPVPosition"].head(n=1).values[0]
        if "DATETIME" in data.columns:
            mpv_line_index = 1
            tmp_data = data.loc[:, self.columns]
        else:
            mpv_line_index = 0
            tmp_data = data

        new_data = {}
        index = 0
        for data_index, line in zip(tmp_data.index, tmp_data.values):  # "DATETIME", "MPVPosition", "CO2_dry", "CH4_dry"
            mpv = float(line[mpv_line_index])
            line[mpv_line_index] = mpv
            line = line.tolist()

            if mpv_line_index == 0:
                line.insert(0, data_index)
            # (*)

            if mpv not in [float(k) for k in self.concentration.keys()]:
                continue
            if line[1] != first_mpv:
                index += 1
                first_mpv = line[1]
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
            df.set_index(columns[0])
            dataframe_dict[index] = df

        return dataframe_dict

    @staticmethod
    def resample_by_1_minute(data):
        """
        Resamples data by 1 minute
        Args:
            data (dict): Data with dict of pd.DataFrame

        Returns:
            dict: Data with dict of pd.DataFrame
        """
        new_data = {}
        for index, array in data.items():
            if array.size:
                df = array.resample(timedelta(minutes=1), on='DATETIME').agg(
                    {'MPVPosition': 'mean',
                     'CH4_dry': ['mean', 'std'],
                     'CO2_dry': ['mean', 'std'],
                     })
                print(f"MPV: '{df['MPVPosition']['mean'].iloc[0]}' Всего строк: {len(df)}")
                new_data[index] = df
        return new_data

    def filter_data_by_std(self, data):
        """
        Applies a data filter by std.
        Keeps at least 3 values with the lowest std value.
        Args:
            data (dict): Data with dict of pd.DataFrame

        Returns:
            dict: Data with dict of pd.DataFrame
        """
        if not self.filter_by_std:
            print("Отключена корректировка по стандартному отклонению.")
        else:
            print(f"Включена корректировка по стандартному отклонению.\n"
                  f"CH4 STD: {self.CH4_std_limit}\n"
                  f"CO2 STD: {self.CO2_std_limit}")
        new_data = {}
        for index, df in data.items():
            if self.filter_by_std:
                std_data = df[
                    df.CH4_dry['std'] < self.CH4_std_limit][
                    df.CO2_dry['std'] < self.CO2_std_limit][
                    pd.notna(df.CH4_dry['std'])][
                    pd.notna(df.CO2_dry['std'])]
                try:
                    df['MPVPosition']['mean'].iloc[0]
                except Exception as err:
                    pprint(df['MPVPosition']['mean'])
                    raise err
                print(f"Цикл: {index}, MPV: '{df['MPVPosition']['mean'].iloc[0]}' Всего строк: {len(df)} -> "
                      f"Минимум корректных строк для CH4_dry и CO2_dry: {len(std_data)}")
                if std_data.size == 0:
                    min_co2 = sorted(df.CO2_dry['std'][pd.notna(df.CO2_dry['std'])])[0]
                    print(f"Минимальное значение CO2 std: {round(min_co2, 4)}")
                    pprint(df)
                    print("======")
                    std_data = df[df.CO2_dry['std'] <= min_co2]
                new_data[index] = std_data
            else:
                print(f"MPV: '{df['MPVPosition']['mean'].iloc[0]}' Всего строк: {len(df)}")
                new_data[index] = df
        return new_data

    def save_to_excel(self, data, filename):
        if isinstance(data, dict):
            data2 = pd.concat(data.values())
        else:
            data2 = data
        try:
            data2.to_excel(filename)
            print(f"Файл создан '{filename}'.")
        except Exception as err:
            input(f"\n[Ошибка] Невозможно сохранить файл '{filename}':\n{err}\n"
                  f"Закройте файл, если он открыт. Чтобы продолжить, нажмите Enter...")
            self.save_to_excel(data, filename)

    @staticmethod
    def read_from_excel(filename=one_minute_resample_filename):
        """
        Asks user to update excel file and then reads it
        Args:
            filename (str): Part of the file name

        Returns:
            pd.DataFrame: or None
        """
        try:
            data = pd.read_excel(filename, header=[0, 1], index_col=0)
            multi_index1 = data.columns
            multi_index2 = data.columns[0]
            while len(multi_index1) != len(multi_index2):
                # input(f"Вы можете изменить промежуточный файл '{filename}'\n"
                #       f"Сохраните его и нажмите Enter, чтобы продолжить...")
                data = pd.read_excel(filename, header=[0, 1], index_col=0)
                multi_index2 = data.columns
                if len(multi_index1) != len(multi_index2):
                    print(f"\n[Ошибка] В файле изменилось количество столбцов с данными. "
                          f"Должно быть ({len(multi_index1) + 1})")
            return data
        except Exception as err:
            print(err)

    def take_last(self, data):
        """
        Takes only a half of the table from the end
        Args:
            data (dict): Data with dict of pd.DataFrame

        Returns:
            dict: Data with dict of pd.DataFrame
        """
        new_data = {}
        for index, array in data.items():
            mpv_to_take_3 = [k for k, v in self.concentration.items() if v['for_calibration'] == 'f']
            if str(array['MPVPosition']['mean'].to_list()[0]) in mpv_to_take_3:
                df = array[-3:]  # flask
            else:
                df = array[-8:]  # balloon
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
            ch4_std = pd.Series(tmp_array['CH4_dry']['mean']).std()
            co2_std = pd.Series(tmp_array['CO2_dry']['mean']).std()
            new_data[index] = tmp_array.mean()
            new_data[index]["datetime"] = array["DATETIME"]["mean"][0]
            new_data[index]["count_mean"] = len(tmp_array)
            new_data[index]["CH4_dry"]['std'] = ch4_std
            new_data[index]["CO2_dry"]['std'] = co2_std
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
        cycle_str = "0"
        for line, date_time in zip(data_dict["data"], data_dict["index"]):
            if self.no_flask:
                if cycle_i < len(self.concentration):
                    cycle_i += 1
                else:
                    cycle_i = 1
                    cycle += 1
                    cycle_str = str(cycle)
            measure_cycles.setdefault(cycle_str, {}).setdefault(
                "data", []).append({str(float(line[0])): [line[1], line[3]]})  # mpv,CH4,CO2

            measure_cycles.setdefault(cycle_str, {}).setdefault(
                "std", []).append({str(float(line[0])): [line[2], line[4]]})  # mpv,CH4_std,CO2_std

            measure_cycles.setdefault(cycle_str, {}).setdefault(
                "count_mean", []).append({str(float(line[0])): line[5]})  # mpv,count_mean

            measure_cycles.setdefault(cycle_str, {}).setdefault(
                "date_time", []).append({str(float(line[0])): date_time})
        mpv_for_calibration = {k: v for k, v in self.concentration.items() if v["for_calibration"] == '1'}

        calibrated_gases = []
        for i, cycle in measure_cycles.items():
            gases_dict = {}
            try:
                for mpv, values in mpv_for_calibration.items():
                    for j, gas in enumerate(["CH4+", "CO2+"]):
                        gases_dict.setdefault(gas, {}).setdefault(
                            "measured", []).extend([d[mpv][j] for d in cycle["data"] if mpv in d.keys()])
                        gases_dict.setdefault(gas, {}).setdefault(
                            "std", []).extend([d[mpv][j] for d in cycle["std"] if mpv in d.keys()])
                        gases_dict.setdefault(gas, {}).setdefault(
                            "assigned", []).append(values[gas])
                        gases_dict.setdefault(gas, {}).setdefault(
                            "date_time", []).extend([d[mpv] for d in cycle["date_time"] if mpv in d.keys()])
                        gases_dict.setdefault(gas, {}).setdefault(
                            "MPV", []).append(str(mpv))
                        gases_dict.setdefault(gas, {}).setdefault(
                            "count_mean", []).extend([d[mpv] for d in cycle["count_mean"] if mpv in d.keys()])
                # concentration CO2(measured CO2)
                co2_coeffs = np.polyfit(gases_dict["CO2+"]["measured"], gases_dict["CO2+"]["assigned"], deg=1)
                # concentration CH4(measured CH4)
                ch4_coeffs = np.polyfit(gases_dict["CH4+"]["measured"], gases_dict["CH4+"]["assigned"], deg=1)
                gases_dict["CH4+"]["coefficients"] = list(ch4_coeffs)
                gases_dict["CO2+"]["coefficients"] = list(co2_coeffs)
                calibrated_gases.append(gases_dict)
            except KeyError as err:
                print(f"\nДля калибровочного газа MPV {err} в цикле №'{i}' измерений не найдено!")
                if calibrated_gases:
                    calibrated_gases.append(calibrated_gases[0])
                    print(f"Будут использованы коэффициенты из предыдущего цикла!\n")
            except Exception as err:
                pprint(gases_dict)
                raise err
        if not calibrated_gases:
            print("Измерений с калибровочными газами не найдено!\n"
                  "(Проверьте соответствие данных с файлом концентраций.)")
            exit(1)

        mpv_not_for_calibration = {k: v for k, v in self.concentration.items() if v["for_calibration"] != '1'}

        gases = []
        for i, cycle in measure_cycles.items():
            gases_dict = {}
            for mpv, values in mpv_not_for_calibration.items():
                balloon_name = values["name"]
                for j, gas in enumerate(["CH4+", "CO2+"]):
                    measured_gas_value = [d[mpv][j] for d in cycle["data"] if mpv in d.keys()]
                    gases_dict.setdefault(balloon_name, {}).setdefault(
                        gas, {})["measured"] = measured_gas_value
                    gases_dict.setdefault(balloon_name, {}).setdefault(
                        gas, {})["std"] = [d[mpv][j] for d in cycle["std"] if mpv in d.keys()]
                    coefficients = calibrated_gases[int(i)][gas]["coefficients"]
                    calculated_value = [coefficients[0] * mgv + coefficients[1] for mgv in measured_gas_value]
                    gases_dict.setdefault(balloon_name, {}).setdefault(
                        gas, {})["calculated"] = calculated_value
                    gases_dict.setdefault(balloon_name, {}).setdefault(
                        gas, {})["date_time"] = [d[mpv] for d in cycle["date_time"] if mpv in d.keys()]
                    gases_dict.setdefault(balloon_name, {}).setdefault(
                        gas, {})["MPV"] = mpv
                    gases_dict.setdefault(balloon_name, {}).setdefault(
                        gas, {})["count_mean"] = [d[mpv] for d in cycle["count_mean"] if mpv in d.keys()]
            gases.append(gases_dict)

        return calibrated_gases, gases

    @staticmethod
    def self_check(cd):
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
                gases = [{} for _, _ in enumerate(calibration_gases_names)]
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
                    for i, _ in enumerate(gas['calculated']):
                        item = {
                            'MPV': gas['MPV'],
                            'calculated': gas['calculated'][i],
                            'count_mean': gas['count_mean'][i],
                            'date_time': gas['date_time'][i],
                            'measured': gas['measured'][i],
                            'std': gas['std'][i],
                        }
                        if name == 'CH4+':
                            ch4_list.append(item)
                        if name == 'CO2+':
                            co2_list.append(item)
        return co2_list, ch4_list

    @staticmethod
    def make_table(data):
        data = pd.DataFrame(data)
        data = data.reindex(['date_time', 'name', 'MPV', 'measured', 'std',
                             'assigned', 'calculated', 'count_mean', 'coefficients'],
                            axis=1)
        data = data.sort_values(by="date_time")
        data = data.reset_index(drop=True)
        data = data[pd.notna(data['date_time'])]
        return data


# class MainApp(tk.Tk):
class MainApp:
    def __init__(self, std=True, no_flask=True):
        # tk.Tk.__init__(self)

        main = Calculate(std, no_flask)
        df = main.get_data()
        print(df)
        exit(0)
        data_dict = main.group_by_mpv_position(df)
        data_dict = main.make_dataframe_dict(data_dict, part=1)
        data_dict = main.resample_by_1_minute(data_dict)

        main.save_to_excel(data_dict, filename=one_minute_resample_filename)
        Chart(self).show(data_dict, mode="data2.index, data2.MPVPosition.mean", line_width=1, marker=None)
        df = main.read_from_excel(filename=one_minute_resample_filename)

        data_dict = main.group_by_mpv_position(df)
        data_dict = main.make_dataframe_dict(data_dict, part=2)
        data_dict = main.filter_data_by_std(data_dict)
        data_dict = main.take_last(data_dict)
        data_dict = main.make_mean(data_dict)
        df = pd.concat(data_dict.values())

        df = pd.DataFrame(dict(MPVPosition=df.MPVPosition["mean"].values,
                               CH4_dry=df.CH4_dry["mean"].values,
                               CH4_dry_std=df.CH4_dry["std"].values,
                               CO2_dry=df.CO2_dry["mean"].values,
                               CO2_dry_std=df.CO2_dry["std"].values,
                               Count_Mean=df.count_mean.values),
                          index=df.datetime.values).sort_index()
        calibrated_data, gases = main.calc_coefficients(df)

        recalculated_calibrated_gases = main.self_check(calibrated_data)

        co2_1, ch4_1 = main.reformat_calibrated_gases(recalculated_calibrated_gases)
        co2_2, ch4_2 = main.reformat_common_gases(gases)
        co2 = co2_1 + co2_2
        ch4 = ch4_1 + ch4_2

        co2_table = main.make_table(co2)
        main.save_to_excel(co2_table, filename=co2_table_filename)

        ch4_table = main.make_table(ch4)
        for column in ["measured", "std", "calculated", "assigned"]:
            ch4_table = main.multiply_1000(ch4_table, column_name=column)
        main.save_to_excel(ch4_table, filename=ch4_table_filename)

        print("\nCO2")
        print(co2_table)
        print("\nCH4")
        print(ch4_table)
        # Chart(self).show(co2_table, mode=["calculated", "measured", "name"])
        # Chart(self).show(ch4_table, mode="measured")


if __name__ == "__main__":
    std = False if 'no_std' in sys.argv else True
    no_flask = True if 'balloon_only' in sys.argv else False
    app = MainApp(std=std, no_flask=no_flask)
    # app.mainloop()
    input("\nДля выхода нажмите Enter...")
