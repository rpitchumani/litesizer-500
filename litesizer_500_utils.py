"""
Utility Functions which work together with the Litesizer 500 Class
"""

from typing import List, Dict, Any
import os
import glob
from pathlib import Path
import datetime
from tqdm import tqdm
import re

from litesizer_500 import Litesizer500


def get_litesizer500_file_list(litesizer500_path: str) -> List:

    litesizer500_psd_pattern = "LS*.xlsx"

    litesizer500_psd_files = glob.glob(os.path.join(litesizer500_path,
                                                    litesizer500_psd_pattern))

    return litesizer500_psd_files


def get_litesizer500_data_from_list(psd_files: List) -> List[Dict]:

    re_pattern_sample_id = "LS(\d{8})_([LR]\d{8}-\d{3})-(\d*)"
    list_litesizer500_data = []

    for psd_file in tqdm(psd_files):

        psd_pathlib = Path(psd_file)
        psd_stem = psd_pathlib.stem

        re_pattern_extract = re.findall(re_pattern_sample_id, psd_stem)

        test_id = re_pattern_extract[0][0]
        file_name_id = re_pattern_extract[0][1]
        run_number = re_pattern_extract[0][2]

        try:
            ls5000 = Litesizer500((psd_file))

            # print(s3500.percentiles)
            # print(file_name_id)

            dict_ls500_data = {
                "test_id": test_id,
                "run_number": run_number,
                "file_name": psd_stem,
                "file_name_id": file_name_id,
                "hydrodynamic_diameter": ls5000.hydrodynamic_diameter,
                "polydispersity_index": ls5000.polydispersity_index,
                "psd_intensity": ls5000.df_psd_intensity_weighted,
                "psd_volume": ls5000.df_psd_volume_weighted,
                "psd_number": ls5000.df_psd_number_weighted
            }

            list_litesizer500_data.append(dict_ls500_data)

        except Exception as e:
            ...

    return list_litesizer500_data
