"""
Class to read exported Anton Paar Litesizer 500 Particle Analyzer XLSX File
2022, Ramanan Pitchumani
"""

from typing import List, Dict, Any
import pandas as pd


class Litesizer500:

    def __init__(self, path_xls: str):

        self.path_xls = path_xls
        self.df = pd.read_excel(
            self.path_xls,
            sheet_name=0,
            header=None,
            skiprows=0,
            engine="openpyxl")

        self.get_sample_information()
        self.get_results()
        self.get_volume_weighted_size_distribution()

    def get_sample_information(self):

        dict_workbook_name = self.get_adjacent_value(
            self.df[[0, 1]], "Workbook name")
        self.workbook_name = list(dict_workbook_name.values())[0]

        dict_measurement_name = self.get_adjacent_value(
            self.df[[0, 1]], "Measurement name")
        self.measurement_name = list(dict_measurement_name.values())[0]

        dict_measurement_mode = self.get_adjacent_value(
            self.df[[0, 1]], "Measurement mode")
        self.measurement_mode = list(dict_measurement_mode.values())[0]

        dict_comment = self.get_adjacent_value(self.df[[0, 1]], "Comment")
        self.comment = list(dict_comment.values())[0]

    def get_results(self):

        dict_hydrodynamic_diameter = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Hydrodynamic diameter")
        self.hydrodynamic_diameter = float(
            list(dict_hydrodynamic_diameter.values())[0])

        dict_polydispersity_index = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Polydispersity index")
        self.polydispersity_index = float(
            list(dict_polydispersity_index.values())[0])

        dict_intercept_g12 = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Intercept g1²")
        self.intercept_g12 = float(
            list(dict_intercept_g12.values())[0])

        dict_baseline = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Baseline")
        self.baseline = float(
            list(dict_baseline.values())[0])

        dict_mean_intensity = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Mean intensity")
        self.mean_intensity = float(
            list(dict_mean_intensity.values())[0])

        dict_absolute_intensity = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Absolute intensity")
        self.absolute_intensity = float(
            list(dict_absolute_intensity.values())[0])

        dict_fit_error = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Fit error")
        self.fit_error = float(
            list(dict_fit_error.values())[0])

        dict_diffusion_coefficient = self.get_adjacent_value(
            self.df[[0, 1, 2, 3]], "Diffusion coefficient")
        self.diffusion_coefficient = float(
            list(dict_diffusion_coefficient.values())[0])

    def get_volume_weighted_size_distribution(self):

        def get_psd(df, weighted):

            list_pos_particle_diameter = self.get_positions_of_value(
                df, "Particle diameter")

            list_pos_weighted = self.get_positions_of_value(
                df, weighted)

            list_col_pos_weighted = [x[1] for x in list_pos_weighted]

            list_col_pos_weighted = \
                [list_pos_particle_diameter[0][1]] + list_col_pos_weighted

            df_psd_weighted = \
                df.iloc[list_pos_particle_diameter[0][0]:,
                        list_col_pos_weighted].dropna(how="all").\
                reset_index(drop=True)

            df_psd_weighted = \
                df_psd_weighted.astype(float, errors="ignore")

            df_psd_weighted = \
                df_psd_weighted.iloc[3:, :].\
                reset_index(drop=True)

            df_psd_weighted.columns = [
                "Particle Diameter (nm)", "Relative Frequency (%)",
                "Cumulative Undersize (%)"]

            return df_psd_weighted

        self.df_psd_intensity_weighted = get_psd(self.df, "Intensity weighted")

        self.df_psd_volume_weighted = get_psd(self.df, "Volume weighted")

        self.df_psd_number_weighted = get_psd(self.df, "Number weighted")

    """
    Utility Functions
    """
    def get_positions_of_value(self,
                               df_object: pd.DataFrame,
                               value: Any) -> List[Dict]:

        """ Get index positions of value in dataframe i.e. dfObj."""
        list_positions = list()

        # Get bool dataframe with True at positions where the given value exists
        result = df_object.isin([value])

        # Get list of columns that contains the value
        series_object = result.any()
        column_names = list(series_object[series_object == True].index)

        # Iterate over list of columns and fetch the rows indexes where value exists
        for col in column_names:

            rows = list(result[col][result[col] == True].index)

            for row in rows:

                list_positions.append((row, col))

        # Return a list of tuples indicating the positions of value in the dataframe
        return list_positions

    def get_adjacent_value(self,
                           df_object: pd.DataFrame,
                           value: Any) -> Any:

        list_position = self.get_positions_of_value(df_object, value)

        result_dict = {}
        if len(list_position) == 1:

            result_dict = {
                df_object.iloc[list_position[0][0], list_position[0][1]]:
                    df_object.iloc[list_position[0][0], list_position[0][1]+1]
            }

            return result_dict

        else:

            return result_dict

    def get_adjacent_value_containing(self,
                                      df_object: pd.DataFrame,
                                      value: Any) -> Any:

        df_contains = df_object.apply(
            lambda col:col.str.contains(value, na=False, case=True), axis=1)

        stacked_positions = df_contains.stack()

        list_position = stacked_positions[stacked_positions].index.tolist()

        result_dict = {}

        if len(list_position) == 1:

            result_dict = {
                df_object.iloc[list_position[0][0], list_position[0][1]]:
                    df_object.iloc[list_position[0][0], list_position[0][1]+1]
            }

            return result_dict

        else:

            return result_dict