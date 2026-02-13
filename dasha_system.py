import logging

import requests
import json
from datetime import datetime, timedelta
logger = logging.getLogger(__name__)
class AstroSystem:
    def __init__(self):
        # දශා සහ නැකත් දත්ත
        self.dasha_info = [
            ("කේතු", 7), ("සිකුරු", 20), ("රවි", 6), ("සඳු", 10),
            ("කුජ", 7), ("රාහු", 18), ("ගුරු", 16), ("ශනි", 19), ("බුධ", 17)
        ]
        self.nak_list = [
            "අස්විද", "බෙරණ", "කැති", "රෙහෙන", "මුවසිරස", "අද", "පුනර්වසු", "පුෂ", "ආස්ලේෂ",
            "මා", "පුවපල්", "උත්තරපල්", "හත", "සිත", "සාති", "විසා", "අනුර", "දෙට",
            "මූල", "පුවසල", "උත්තරසල", "සුවණ", "දෙණට", "සියාවස", "පුවපුටුප", "උත්තරපුටුප", "රේවතී"
        ]

    def fetch_moon_degree(self, utc_dt):
        """NASA API එකෙන් අදාළ වේලාවට සඳුගේ අංශකය ලබාගනී"""
        try:

            # දැන් අපි සෘජුවම utc_dt භාවිතා කරමු
            start_str = utc_dt.strftime('%Y-%m-%d %H:%M')
            end_str = (utc_dt + timedelta(minutes=1)).strftime('%Y-%m-%d %H:%M')
            # 2. NASA API Call
            url = "https://ssd.jpl.nasa.gov/api/horizons.api"
            params = {
                "format": "json",
                "COMMAND": "'301'", "OBJ_DATA": "NO", "MAKE_EPHEM": "YES",
                "EPHEM_TYPE": "OBSERVER", "CENTER": "'500@399'",
                "QUANTITIES": "'1'", "START_TIME": f"'{start_str}'",
                "STOP_TIME": f"'{end_str}'", "STEP_SIZE": "'1m'", "ANG_FORMAT": "DEG"
            }


            response = requests.get(url, params=params)
            data = response.json()

            if "result" in data:
                raw_output = data["result"]
                # $$SOE සහ $$EOE අතර දත්ත සෙවීම
                soe_idx = raw_output.find("$$SOE") + 5
                eoe_idx = raw_output.find("$$EOE")
                data_line = raw_output[soe_idx:eoe_idx].strip().split('\n')[0]
                # 3 වැනි තීරුවේ ඇත්තේ අංශකයයි (RA)
                return float(data_line.split()[2])
        except Exception as e:
            print(f"NASA API Error: {e}")
            return None

    def get_nasa_analysis(self, birth_date_str, birth_time_str):
        # 1. මුලින්ම String එක DateTime Object එකක් බවට පත් කරගන්න
        dt_obj = datetime.strptime(f"{birth_date_str} {birth_time_str}", "%Y/%m/%d %H:%M")

        # 2. ස්වයංක්‍රීය කාල නිවැරදි කිරීම (Timezone Correction)
        # 1996-05-25 සිට 2006-04-15 දක්වා කාලය තුළ ලංකාව GMT+6 ලෙස ක්‍රියා කළා
        tz_start = datetime(1996, 5, 25)
        tz_end = datetime(2006, 4, 15)

        if tz_start <= dt_obj <= tz_end:
            # 2004 වැනි වසරවල් සඳහා පැය 6ක් අඩු කරයි
            utc_dt = dt_obj - timedelta(hours=6)
        else:
            # අනෙක් සෑම කාලයකටම පැය 5:30ක් අඩු කරයි
            utc_dt = dt_obj - timedelta(hours=5, minutes=30)

        # 3. සඳුගේ අංශකය ලබාගැනීම (දැන් අපි utc_dt යොදාගන්නවා)
        # මීට පෙර ඔබ birth_date_str වැනි strings යැවූ නිසා fetch_moon_degree එකත් මීට ගැලපෙන ලෙස වෙනස් විය යුතුයි
        logger.info("Start NASA fetch")
        moon_deg = self.fetch_moon_degree(utc_dt)
        logger.info("End NASA fetch")
        if moon_deg is None: return {"error": "NASA data fetch failed"}

        # 4. නිරයන අංශකය සහ දශා ගණනය (පැරණි logic එකම භාවිතා වේ)
        ayanamsa = 23.9
        nirayana_lon = (moon_deg - ayanamsa) % 360
        nak_pos_raw = nirayana_lon / (360/27)
        dasha_idx = int(nak_pos_raw % 9)
        remaining_ratio = 1 - (nak_pos_raw % 1)

        d_name, d_total_y = self.dasha_info[dasha_idx]
        rem_years = d_total_y * remaining_ratio

        # වර්තමාන කාලය (2026-02-12)
        today_dt = datetime(2026, 2, 12)

        # මහ දශා සොයාගැනීමේ Logic
        current_marker = dt_obj + timedelta(days=rem_years * 365.25)
        active_idx = dasha_idx
        maha_start_date = dt_obj

        if today_dt > current_marker:
            while current_marker < today_dt:
                active_idx = (active_idx + 1) % 9
                maha_start_date = current_marker
                current_marker += timedelta(days=self.dasha_info[active_idx][1] * 365.25)
        else:
            maha_start_date = dt_obj

        # අන්තර් දශා ලැයිස්තුව සැකසීම
        m_planet, m_years = self.dasha_info[active_idx]
        antar_list = []
        a_pointer = maha_start_date
        for j in range(9):
            a_idx = (active_idx + j) % 9
            a_planet, a_years = self.dasha_info[a_idx]
            duration = (m_years * a_years / 120) * 365.25
            a_end = a_pointer + timedelta(days=duration)
            status = "අනාගතයේ"
            if a_end < today_dt: status = "ගෙවී අවසන්"
            elif a_pointer <= today_dt <= a_end: status = "වර්තමානයේ ගෙවේ"
            antar_list.append({"ග්‍රහයා": a_planet, "කාලය": f"{a_pointer.strftime('%Y-%m-%d')} සිට {a_end.strftime('%Y-%m-%d')}", "තත්ත්වය": status})
            a_pointer = a_end

        return {
            "නැකත": self.nak_list[int(nirayana_lon / (360/27))],
            "නිරයන_අංශකය": f"{nirayana_lon:.4f}",
            "වර්තමාන_දශා_විස්තර": {
                "මහ_දශාව": m_planet,
                "අන්තර්_දශා_ලැයිස්තුව": antar_list
            },
            "අනාගත_දශා_විස්තර": { "මීළඟ_මහ_දශාව": self.dasha_info[(active_idx + 1) % 9][0], "ආරම්භය": current_marker.strftime('%Y-%m-%d') }
        }

    def print_raw_nasa_json(self, nasa_data):
        print("\n--- RAW NASA JSON OUTPUT ---")
        print(json.dumps(nasa_data, indent=4, ensure_ascii=False))

if __name__ == "__main__":
    app = AstroSystem()
    # Hard-coded Test Data
    test_dob = "2007/08/25"
    test_time = "08:14"

    nasa_output = app.get_nasa_analysis(test_dob, test_time)
    app.print_raw_nasa_json(nasa_output)