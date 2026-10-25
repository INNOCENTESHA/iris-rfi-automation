"""Workbook routing shared by the IRIS scripts.

65.5 keeps RFIs and PHYSICAL PROGRESS in one Detailed Log workbook.
163.22 retains its existing progress workbook until its tab is migrated.
"""
import os

RFI_ID = {
    "65.5": "1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU",
    "163.22": "1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA",
}
RFI_NAMES = {
    "65.5": "65.5 IRIS RFI Detailed Log",
    "163.22": "163.22 IRIS RFI Detailed Log",
}
RFI_ID_BY_NAME = {RFI_NAMES[project]: ssid for project, ssid in RFI_ID.items()}
PP_SHEET_NAME = "PHYSICAL PROGRESS"
PP_ID = {
    "65.5": RFI_ID["65.5"],
    "163.22": os.environ.get("IRIS_PP_ID_163_22", "").strip()
              or "1YbDCzePl76CqtO9dLl2YZjIZtVmJKSgVkmU6LKqXKUc",
}


def physical_progress_tab(titles):
    """Require the intended tab; never fall back to an RFI data tab."""
    if PP_SHEET_NAME not in titles:
        raise ValueError(f"Required tab {PP_SHEET_NAME!r} is missing")
    return PP_SHEET_NAME
