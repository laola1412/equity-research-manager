from PyQt6.QtWidgets import QStyledItemDelegate
import datetime


class DollarSignFormat(QStyledItemDelegate):
    def displayText(self, value, locale):
        try:
            if value is not None and len(str(value)) > 0:
                return f"${value}"
            else:
                return ""
        except:
            print("Error formatting dollar sign.")
            return ""


class TitleCaseFormat(QStyledItemDelegate):
    def displayText(self, value, locale):
        if value is not None and len(str(value)) > 0:
            return f"{value.title()}"
        else:
            return ""


class PercentFormat(QStyledItemDelegate):
    def displayText(self, value, locale):
        if value is not None and len(str(value)) > 0:
            return f"{value}%"
        else:
            return ""


class DateFormat(QStyledItemDelegate):
    def displayText(self, value, locale):
        if value is not None and len(str(value)) > 0:
            date_object = datetime.datetime.strptime(value, "%Y%m%d")
            formatted_date = date_object.strftime("%d.%m.%Y")
            return f"{formatted_date}"
        else:
            return ""
