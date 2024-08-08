from os import getcwd
import os
import sys
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as fd
import tkinter.messagebox as mb

from tkcalendar import Calendar

import report_emk
import report_bunk_50
import report_phone_adress
import report_operations
import report_services
from validators import ValidateError
import loading_window

__version__ = "2.1.0"
__author__ = "DinoWithPython"
__copyright__ = "2024, ГБУЗ Мытищинская ОКБ"
__contact__ = "<email: pa.dmi@rambler.ru>"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class Pathfile:
    """Объект для сохранения пути."""

    path = None


class Emk(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.filepath_emk = Pathfile()

        text_file_emk = tk.Label(
            container.window_emk,
            text="Для создания отчета по ЭМК, выберите файл",
            font=("Microsoft Sans Serif", 16),
        )

        btn_file_emk = tk.Button(
            container.window_emk,
            text="Выбрать файл...",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: container.choose_file(btn_file_emk, self.filepath_emk),
        )
        text_file_emk.place(x=10, y=10)
        btn_file_emk.place(x=580, y=10)
        self.btn_start_emk = tk.Button(
            container.window_emk,
            text="Сформировать",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: self.read_and_create_summary_emk(container),
        )
        self.btn_start_emk.place(x=780, y=10)
        self.need_pdo = tk.BooleanVar()
        self.need_pdo.set(0)
        need_pdo = tk.Checkbutton(
            container.window_emk,
            text="↑ включая ПДО",
            font=("Microsoft Sans Serif", 16),
            variable=self.need_pdo,
            onvalue=1,
            offvalue=0,
        )
        need_pdo.place(x=780, y=58)
        self.btn_info_emk = tk.Button(
            container.window_emk,
            text="Инфо",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=self.helping_emk,
        )
        self.btn_info_emk.place(x=780, y=110)

        self.text_date_emk = tk.Label(
            container.window_emk,
            text="Отчет на дату: <Дата не выбрана>",
            font=("Microsoft Sans Serif", 16),
        )

        self.text_date_emk.place(x=10, y=50)

        self.clear_date_emk = tk.Button(
            container.window_emk,
            text="Очистить ->",
            font=("Microsoft Sans Serif", 10),
            command=lambda: self.del_date(container),
        )
        self.clear_date_emk.place(x=480, y=60)

        self.date_button_emk = tk.Button(
            container.window_emk,
            text="Выбрать дату",
            font=("Microsoft Sans Serif", 16),
            command=lambda: self.open_calendar(container),
        )
        self.date_button_emk.place(x=580, y=60)
        frame_reports = tk.LabelFrame(
            container.window_emk,
            text="Доп. отчеты")
        frame_reports.place(x=780, y=160)
        self.need_lis = tk.BooleanVar()
        self.need_lis.set(1)
        button_lis = tk.Checkbutton(
            frame_reports,
            text="Отчет по ЛИС",
            font=("Microsoft Sans Serif", 16),
            variable=self.need_lis,
            onvalue=1,
            offvalue=0,
        )
        button_lis.pack(anchor="w")
        self.need_instrumental = tk.BooleanVar()
        self.need_instrumental.set(1)
        button_need_instrumental = tk.Checkbutton(
            frame_reports,
            text="Отчет по Инстр.",
            font=("Microsoft Sans Serif", 16),
            variable=self.need_instrumental,
            onvalue=1,
            offvalue=0,
        )
        button_need_instrumental.pack(anchor="w")
        self.need_cons = tk.BooleanVar()
        self.need_cons.set(1)
        button_need_cons = tk.Checkbutton(
            frame_reports,
            text="Отчет по Конс.",
            font=("Microsoft Sans Serif", 16),
            variable=self.need_cons,
            onvalue=1,
            offvalue=0,
        )
        button_need_cons.pack(anchor="w")

    def helping_emk(self):
        link = "http://bi.mz.mosreg.ru/#form/oformlen_emk_al_23f"
        msg = """
        Отчет с BI "Оформление электронной медицинской карты (п)".

        Доступен по ссылке: {}

        Отчеты по ЛИС и Инструментальным исследованиям формируются по всем датам в файле,
        без привязки к выбранной дате.

        Скопировать ссылку в буфер обмена?
        """.format(
            link
        )

        help_window = tk.Tk()
        help_window.title("Описание отчета по ЭМК")
        help_window.iconbitmap(resource_path("images/icon.ico"))
        help_window.configure(background="WHITE")
        text = tk.Label(help_window,
                        text=msg,
                        font=("", 14),
                        justify="left")
        text.pack(expand=True)
        self.button_help_emk = tk.Button(
            help_window,
            text="Скопировать ссылку",
            font=("", 14),
            command=lambda: self.copy_link(frame=help_window, link=link),
        )
        self.button_help_emk.pack(anchor="se", padx=10, pady=10)

    def copy_link(self, frame, link):
        """Копирует ссылку отчета из окна Инфо."""
        frame.clipboard_clear()
        frame.clipboard_append(link)
        self.button_help_emk["text"] = "Скопирована"
        self.button_help_emk["bg"] = "LimeGreen"

    def delete_panel_errors(self):
        """Уничтожает панели с сошибками."""
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        try:
            self.troubles_excel.destroy()
        except AttributeError:
            pass
        try:
            self.file_processing.destroy()
        except AttributeError:
            pass

    def open_calendar(self, container):
        """Функция открывает календарь."""
        try:
            container.text_error.destroy()
        except AttributeError:
            pass
        try:
            # Если была ошибка при открытии файла эксель
            self.troubles_excel.destroy()
        except AttributeError:
            pass
        try:
            self.file_processing.destroy()
        except AttributeError:
            pass
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        self.calend = tk.Toplevel(self)

        self.cal = Calendar(self.calend, font="Arial 14")
        self.cal.pack(fill="both", expand=True)
        ttk.Button(self.calend, text="Выбрать", command=self.check_date).pack()

    def check_date(self):
        """Меняет лейбл, записывая туда дату."""
        date = self.cal.selection_get().strftime("%d.%m.%Y")
        self.text_date_emk.config(text=f"Отчет на дату: {date}")
        self.calend.destroy()
        self.date_button_emk["bg"] = "LimeGreen"
        self.date_button_emk["text"] = "Выбрана"
        self.update()

    def del_date(self, container):
        """Очищает поле даты."""
        self.text_date_emk.config(text="Отчет на дату: <Дата не выбрана>")
        self.date_button_emk["bg"] = container.cget("bg")
        self.date_button_emk["text"] = "Выбрать дату"
        self.text_date_emk.update()

    def read_and_create_summary_emk(self, container):
        """Формирует отчет по отсутсвию осмотров и детально по заполнению КВС."""
        self.delete_panel_errors()
        self.file_processing = tk.Label(
            container.window_emk,
            font=("Microsoft Sans Serif", 16),
            text="Файл обрабатывается, пожалуйста, подождите...",
        )
        self.file_processing.place(x=10, y=110)
        self.update()

        try:
            if self.filepath_emk.path is None:
                raise ValidateError("Выберите файл с отчетом!")
            need_pdo = self.need_pdo.get()
            if need_pdo:
                excel_document = report_emk.EmkReport(self.filepath_emk.path, True)
            else:
                excel_document = report_emk.EmkReport(self.filepath_emk.path)
            data_excel = excel_document.open_file_return_data()
            get_data = self.text_date_emk.cget("text").split(": ")[1]
            if get_data != "<Дата не выбрана>":
                input_date = get_data
                datas = excel_document.processing_report(data_excel, input_date)
                excel_document.save_files(*datas)
            else:
                datas = excel_document.processing_report(data_excel)
                excel_document.save_files(*datas)
            need_lis = self.need_lis.get()
            if need_lis:
                lis = report_emk.LisIdentificator(report_emk.HEADINGS[17], report_emk.HEADINGS[18])
                data = lis.processing()
                lis.save_file(data,"Отчет по ЛИС в ЭМК", "ЛИС по выписанным пациентам", "ЛИС")
            need_inst = self.need_instrumental.get()
            if need_inst:
                ins = report_emk.InstIdentificator(report_emk.HEADINGS[19], report_emk.HEADINGS[20])
                data_ins = ins.processing()
                ins.save_file(data_ins, "Отчет по Инстр.напр. в ЭМК", "Инструментальная диагностика по выписанным пациентам", "Инструм на")
            need_cons = self.need_cons.get()
            if need_cons:
                cons = report_emk.ConsIdentificator(report_emk.HEADINGS[21], report_emk.HEADINGS[22])
                data_cons = cons.processing()
                cons.save_file(data_cons, "Отчет по Конс. в ЭМК", "Оформление консультативных услуг по выписанным пациентам", "Консультации на")
            
        except Exception as e:
            print(type(e))
            if isinstance(e, ValueError) or isinstance(e, ValidateError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка валидации файла!"
            else:
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = f"Неизвестная ошибка. {type(e)}"
            self.error_panel = tk.Text(
                container.window_emk, bg="#FFCCCC", width=80, height=7
            )
            self.error_panel.insert(tk.INSERT, str(e))
            self.error_panel.place(x=10, y=180)
            return
        else:
            self.file_processing[
                "text"
            ] = "Файл успешно обработан и сохранён в текущей папке."
            self.file_processing["fg"] = "LimeGreen"
            self.update()


class Bunk(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.filepath_bunks = Pathfile()
        text_file_bunks = tk.Label(
            container.window_bunks,
            text="Для создания отчета по Койкам и 50+, выберите файл",
            font=("Microsoft Sans Serif", 16),
        )

        btn_file_bunks = tk.Button(
            container.window_bunks,
            text="Выбрать файл...",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: container.choose_file(btn_file_bunks, self.filepath_bunks),
        )
        text_file_bunks.place(x=10, y=10)
        btn_file_bunks.place(x=600, y=10)
        self.btn_start_bunks = tk.Button(
            container.window_bunks,
            text="Сформировать",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: self.read_and_create_summary_bunks(container),
        )
        self.btn_start_bunks.place(x=800, y=10)
        self.btn_info_bunks = tk.Button(
            container.window_bunks,
            text="Инфо",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=self.helping_bunks,
        )
        self.btn_info_bunks.place(x=800, y=60)

        self.count_days = tk.Label(
            container.window_bunks,
            text="Изменить количество дней(по умолчанию 50)",
            font=("Microsoft Sans Serif", 16),
        )
        self.count_days.place(x=10, y=50)

        var = tk.IntVar()
        var.set(report_bunk_50.COUNT_DAYS)
        check = (self.register(self.validate_days), "%P")
        self.spin_50 = tk.Spinbox(
            container.window_bunks,
            from_=1,
            to=100,
            width=3,
            textvariable=var,
            validate="key",
            validatecommand=check,
            font=("Microsoft Sans Serif", 16),
        )
        self.spin_50.place(x=600, y=60)

    def file_not_found_bunks(self):
        msg = """
        Программа не нашла информацию о коечном фонде. 
        В директории с файлом .exe создан документ "Отделения и койки.xlsx".
        Заполните информацию в нём об отделениях и количестве коек по аналогии с примером.
        """
        mb.showinfo("Необходима информация!", msg)

    def helping_bunks(self):
        report = "Список пациентов, находящихся на лечении"
        msg = """
        Отчет из ЕМИАС Стационар "{}".

        Учитывает всех пациентов в отчете, без учета исхода. Желательно брать период "сегодня".

        Информацию по койкам программа получает из файла "Отделения и койки.xlsx"
        Если файл есть в папке с программой, то информация подтянется автоматически.
        Если файла с коечным фондом нет, то программа сохранит образец для заполнения в папке с .exe.

        Наименование отделения необходимо указывать как в отчете.

        Сформировать образец сейчас(заменит файл, если он есть в текущей папке)?
        """.format(report)

        help_window = tk.Tk()
        help_window.title("Описание отчета по Койкам")
        help_window.iconbitmap(resource_path("images/icon.ico"))
        help_window.configure(background="WHITE")
        text = tk.Label(help_window,
                        text=msg,
                        font=("", 14),
                        justify="left")
        text.pack(expand=True)
        self.button_help_bunk = tk.Button(
            help_window,
            text="Скопировать название отчета",
            font=("", 14),
            command=lambda: self.copy_link(frame=help_window, link=report),
        )
        self.button_help_bunk.pack(anchor="se", padx=10, pady=10)
        self.create_sample_btn = tk.Button(
            help_window,
            text="Создать образец",
            font=("", 14),
            command=lambda: self.create_smple(),
        )
        self.create_sample_btn.pack(anchor="se", padx=10, pady=10)

    def copy_link(self, frame, link):
        """Копирует ссылку отчета из окна Инфо."""
        frame.clipboard_clear()
        frame.clipboard_append(link)
        self.button_help_bunk["text"] = "Скопировано"
        self.button_help_bunk["bg"] = "LimeGreen"

    def create_smple(self):
        """Создает образец отчета."""
        report_bunk_50.BunkReport.create_sample()
        self.create_sample_btn["text"] = "Создан"
        self.create_sample_btn["bg"] = "LimeGreen"

    def delete_panel_errors(self):
        """Уничтожает панели с сошибками."""
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        try:
            self.file_processing.destroy()
        except AttributeError:
            pass

    def validate_days(self, text):
        """Валидация значения в спинбоксе для 50+."""
        return text.isdigit()

    def read_and_create_summary_bunks(self, container):
        """Формирует отчет по разнице коек и пациентов более 50 дней."""
        self.delete_panel_errors()
        self.file_processing = tk.Label(
            container.window_bunks,
            font=("Microsoft Sans Serif", 16),
            text="Файл обрабатывается, пожалуйста, подождите...",
        )
        self.file_processing.place(x=10, y=110)
        self.update()

        try:
            if self.filepath_bunks.path is None:
                raise ValidateError("Выберите файл с отчетом!")
            bunk = report_bunk_50.BunkReport(self.filepath_bunks.path)
            days = int(self.spin_50.get())
            report_bunk_50.COUNT_DAYS = days
            data_bunks, data_50 = bunk.open_file_return_data()
            data_after_processing = bunk.processing(data_bunks, data_50)
            bunk.save_in_files(*data_after_processing)
        except Exception as e:
            print(type(e))
            if isinstance(e, FileNotFoundError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = 'Заполните файл "Отделения и койки.xlsx!"'
            elif isinstance(e, ValueError) or isinstance(e, ValidateError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка валидации файла!"
            else:
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = f"Неизвестная ошибка. {type(e)}"
            self.error_panel = tk.Text(
                container.window_bunks, bg="#FFCCCC", width=80, height=7
            )
            self.error_panel.insert(tk.INSERT, str(e))
            self.error_panel.place(x=10, y=180)
            return
        else:
            self.file_processing[
                "text"
            ] = "Файл успешно обработан и сохранён в текущей папке."
            self.file_processing["fg"] = "LimeGreen"
            self.update()


class Phone(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.filepath_phone = Pathfile()

        text_file_phone = tk.Label(
            container.window_phone,
            text="Для создания отчета по Тел. и адресу, выберите файл",
            font=("Microsoft Sans Serif", 16),
        )

        btn_file_phone = tk.Button(
            container.window_phone,
            text="Выбрать файл...",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: container.choose_file(btn_file_phone, self.filepath_phone),
        )
        text_file_phone.place(x=10, y=10)
        btn_file_phone.place(x=600, y=10)
        self.btn_start_phone = tk.Button(
            container.window_phone,
            text="Сформировать",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: self.read_and_create_summary_phone(container),
        )
        self.btn_start_phone.place(x=800, y=10)
        self.btn_info_phone = tk.Button(
            container.window_phone,
            text="Инфо",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=self.helping_phone,
        )
        self.btn_info_phone.place(x=800, y=60)

    def helping_phone(self):
        report = "Список поступивших пациентов по дате и времени"
        msg = """
        Отчет из ЕМИАС Стационар "{}".
        Скопировать название отчета в буфер?
        """.format(
            report
        )
        help_window = tk.Tk()
        help_window.title("Описание отчета по Телефонам и адресам")
        help_window.iconbitmap(resource_path("images/icon.ico"))
        help_window.configure(background="WHITE")
        text = tk.Label(help_window,
                        text=msg,
                        font=("", 14),
                        justify="left")
        text.pack(expand=True)
        self.button_help_phone = tk.Button(
            help_window,
            text="Скопировать название",
            font=("", 14),
            command=lambda: self.copy_link(frame=help_window, report=report),
        )
        self.button_help_phone.pack(anchor="se", padx=10, pady=10)

    def copy_link(self, frame, report):
        """Копирует ссылку отчета из окна Инфо."""
        frame.clipboard_clear()
        frame.clipboard_append(report)
        self.button_help_phone["text"] = "Скопировано"
        self.button_help_phone["bg"] = "LimeGreen"

    def delete_panel_errors(self):
        """Уничтожает панели с сошибками."""
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        try:
            self.troubles_excel.destroy()
        except AttributeError:
            pass
        try:
            self.file_processing.destroy()
        except AttributeError:
            pass

    def read_and_create_summary_phone(self, container):
        """Формирует отчет по отсутсвию телефона и адреса."""
        self.delete_panel_errors()
        self.file_processing = tk.Label(
            container.window_phone,
            font=("Microsoft Sans Serif", 16),
            text="Файл обрабатывается, пожалуйста, подождите...",
        )
        self.file_processing.place(x=10, y=110)
        self.update()
        try:
            if self.filepath_phone.path is None:
                raise ValidateError("Выберите файл с отчетом!")
            excel_document = report_phone_adress.PhoneReport(self.filepath_phone.path)
            data_phone_excel, data_adress_excel = excel_document.open_file_return_data()
            excel_document.processing_and_save(data_phone_excel, data_adress_excel)
        except Exception as e:
            print(type(e))
            if isinstance(e, ValueError) or isinstance(e, ValidateError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка валидации файла!"
            elif isinstance(e, TypeError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка при открытии файла!"
            else:
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = f"Неизвестная ошибка. {type(e)}"
            self.error_panel = tk.Text(
                container.window_phone, bg="#FFCCCC", width=80, height=7
            )
            self.error_panel.insert(tk.INSERT, str(e))
            self.error_panel.place(x=10, y=180)
            return
        else:
            self.file_processing[
                "text"
            ] = "Файл успешно обработан и сохранён в текущей папке."
            self.file_processing["fg"] = "LimeGreen"
            self.update()


class Operation(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)
        self.filepath_operation = Pathfile()

        text_file_operation = tk.Label(
            container.window_operation,
            text="Для создания отчета по Операциям, выберите файл",
            font=("Microsoft Sans Serif", 16),
        )

        btn_file_operation = tk.Button(
            container.window_operation,
            text="Выбрать файл...",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: container.choose_file(
                btn_file_operation, self.filepath_operation
            ),
        )
        text_file_operation.place(x=10, y=10)
        btn_file_operation.place(x=600, y=10)
        self.btn_start_operation = tk.Button(
            container.window_operation,
            text="Сформировать",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: self.read_and_create_summary_operation(container),
        )
        self.btn_start_operation.place(x=800, y=10)
        self.btn_info_operation = tk.Button(
            container.window_operation,
            text="Инфо",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=self.helping_operation,
        )
        self.btn_info_operation.place(x=800, y=60)

        self.only_a16 = tk.BooleanVar()
        self.only_a16.set(1)
        only_a16 = tk.Checkbutton(
            container.window_operation,
            text="Только А16",
            font=("Microsoft Sans Serif", 16),
            variable=self.only_a16,
            onvalue=1,
            offvalue=0,
        )
        only_a16.place(x=780, y=110)

    def helping_operation(self):
        report = "Список пациентов по отделениям, с указанием реанимационных периодов и операций"
        msg = """
        Отчет из ЕМИАС Стационар "{}".

        Игнорирует наименования операций включающие "вентиляц".

        Скопировать название отчета в буфер?
        """.format(
            report
        )
        help_window = tk.Tk()
        help_window.title("Описание отчета по Операциям")
        help_window.iconbitmap(resource_path("images/icon.ico"))
        help_window.configure(background="WHITE")
        text = tk.Label(help_window,
                        text=msg,
                        font=("", 14),
                        justify="left")
        text.pack(expand=True)
        self.button_help_oper = tk.Button(
            help_window,
            text="Скопировать название",
            font=("", 14),
            command=lambda: self.copy_link(frame=help_window, report=report),
        )
        self.button_help_oper.pack(anchor="se", padx=10, pady=10)

    def copy_link(self, frame, report):
        """Копирует ссылку отчета из окна Инфо."""
        frame.clipboard_clear()
        frame.clipboard_append(report)
        self.button_help_oper["text"] = "Скопировано"
        self.button_help_oper["bg"] = "LimeGreen"

    def delete_panel_errors(self):
        """Уничтожает панели с сошибками."""
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        try:
            self.troubles_excel.destroy()
        except AttributeError:
            pass
        try:
            self.file_processing.destroy()
        except AttributeError:
            pass

    def read_and_create_summary_operation(self, container):
        """Формирует отчет по отсутсвию телефона и адреса."""
        self.delete_panel_errors()
        self.file_processing = tk.Label(
            container.window_operation,
            font=("Microsoft Sans Serif", 16),
            text="Файл обрабатывается, пожалуйста, подождите...",
        )
        self.file_processing.place(x=10, y=110)
        self.update()

        try:
            if self.filepath_operation.path is None:
                raise ValidateError("Выберите файл с отчетом!")
            excel_document = report_operations.OperationReport(self.filepath_operation.path)
            only_a16 = self.only_a16.get()
            data_excel = excel_document.open_file_return_data(only_a16)
            excel_document.processing_and_save(data_excel)
        except Exception as e:
            print(type(e))
            if isinstance(e, ValueError) or isinstance(e, ValidateError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка валидации файла!"
            elif isinstance(e, TypeError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка при открытии файла!"
            else:
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = f"Неизвестная ошибка. {type(e)}"
            self.error_panel = tk.Text(
                container.window_operation, bg="#FFCCCC", width=80, height=7
            )
            self.error_panel.insert(tk.INSERT, str(e))
            self.error_panel.place(x=10, y=180)
            return
        else:
            self.file_processing[
                "text"
            ] = "Файл успешно обработан и сохранён в текущей папке."
            self.file_processing["fg"] = "LimeGreen"
            self.update()


class Service(ttk.Frame):
    """Отчет по услугам ЛИС и инст."""
    def __init__(self, container):
        super().__init__(container)
        self.filepath_services = Pathfile()

        text_file_services = tk.Label(
            container.window_services,
            text="Для создания отчета по Напр. на услуги,выберите файл",
            font=("Microsoft Sans Serif", 16),
        )

        btn_file_services = tk.Button(
            container.window_services,
            text="Выбрать файл...",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: container.choose_file(
                btn_file_services, self.filepath_services
            ),
        )
        text_file_services.place(x=10, y=10)
        btn_file_services.place(x=600, y=10)
        self.btn_start_services= tk.Button(
            container.window_services,
            text="Сформировать",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=lambda: self.read_and_create_summary_operation(container),
        )
        self.btn_start_services.place(x=800, y=10)
        self.btn_info_services = tk.Button(
            container.window_services,
            text="Инфо",
            font=("Microsoft Sans Serif", 16),
            pady=2,
            command=self.helping_services,
        )
        self.btn_info_services.place(x=800, y=60)

    def helping_services(self):
        report = "Список пациентов, которым выданы направления на услуги"
        msg = """
        Для документа по ЛИС формируется из кодов: "A08", "A09", "A12", "A26", "B03"

        Отчет из ЕМИАС Стационар "{}".

        Скопировать название отчета в буфер?
        """.format(
            report
        )
        help_window = tk.Tk()
        help_window.title("Описание отчета по Услугам")
        help_window.iconbitmap(resource_path("images/icon.ico"))
        help_window.configure(background="WHITE")
        text = tk.Label(help_window,
                        text=msg,
                        font=("", 14),
                        justify="left")
        text.pack(expand=True)
        self.button_help_ser = tk.Button(
            help_window,
            text="Скопировать название",
            font=("", 14),
            command=lambda: self.copy_link(frame=help_window, report=report),
        )
        self.button_help_ser.pack(anchor="se", padx=10, pady=10)

    def copy_link(self, frame, report):
        """Копирует ссылку отчета из окна Инфо."""
        frame.clipboard_clear()
        frame.clipboard_append(report)
        self.button_help_ser["text"] = "Скопировано"
        self.button_help_ser["bg"] = "LimeGreen"

    def delete_panel_errors(self):
        """Уничтожает панели с сошибками."""
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        try:
            self.troubles_excel.destroy()
        except AttributeError:
            pass
        try:
            self.file_processing.destroy()
        except AttributeError:
            pass

    def read_and_create_summary_operation(self, container):
        """Формирует отчет по отсутсвию телефона и адреса."""
        self.delete_panel_errors()
        self.file_processing = tk.Label(
            container.window_services,
            font=("Microsoft Sans Serif", 16),
            text="Файл обрабатывается, пожалуйста, подождите...",
        )
        self.file_processing.place(x=10, y=110)
        self.update()

        try:
            if self.filepath_services.path is None:
                raise ValidateError("Выберите файл с отчетом!")
            excel_document = report_services.ServicesReport(self.filepath_services.path)
            inst_from_excel, lis_from_excel, svod_from_excel = excel_document.open_file_return_data()
            excel_document.save_files(inst_from_excel, lis_from_excel, svod_from_excel)
        except Exception as e:
            print(type(e))
            if isinstance(e, ValueError) or isinstance(e, ValidateError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка валидации файла!"
            elif isinstance(e, TypeError):
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = "Ошибка при открытии файла!"
            else:
                self.file_processing["fg"] = "Crimson"
                self.file_processing["text"] = f"Неизвестная ошибка. {type(e)}"
            self.error_panel = tk.Text(
                container.window_services, bg="#FFCCCC", width=80, height=7
            )
            self.error_panel.insert(tk.INSERT, str(e))
            self.error_panel.place(x=10, y=180)
            return
        else:
            self.file_processing[
                "text"
            ] = "Файл успешно обработан и сохранён в текущей папке."
            self.file_processing["fg"] = "LimeGreen"
            self.update()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Обработка ЭМК")
        self.geometry("1000x400")
        self.resizable(width=False, height=False)
        self.tabs = ttk.Notebook(self)
        self.iconbitmap(resource_path("images/icon.ico"))

        self.window_emk = ttk.Frame(self.tabs)
        self.window_bunks = ttk.Frame(self.tabs)
        self.window_phone = ttk.Frame(self.tabs)
        self.window_operation = ttk.Frame(self.tabs)
        self.window_services = ttk.Frame(self.tabs)
        self.tabs.add(self.window_emk, text="ЭМК")
        self.tabs.add(self.window_bunks, text="Койки и 50+")
        self.tabs.add(self.window_phone, text="Телефоны и адреса")
        self.tabs.add(self.window_operation, text="Операции")
        self.tabs.add(self.window_services, text="Напр.услуги")

        self.creator_text = tk.Label(
            self,
            text=f"Разработано ОВCиПИС ГБУЗ МО Мытищиская ГКБ 2023г. | v{__version__}",
            font=("Arial", 8),
            fg="grey",
        )
        self.creator_text.place(x=10, y=370)
        self.tabs.pack(expand=1, fill="both")

    def delete_panel_errors(self):
        """Уничтожает панели с сошибками."""
        try:
            self.error_panel.destroy()
        except AttributeError:
            pass
        try:
            self.troubles_excel.destroy()
        except AttributeError:
            pass

    def choose_file(self, btn_file, file):
        try:
            self.file_processing.destroy()
            btn_file["bg"] = self.cget("bg")
        except AttributeError:
            pass
        self.delete_panel_errors()

        filetypes = (("Microsoft excel 2007/2010", "*.xlsx"), ("Любой", "*"))
        initialdir = getcwd()
        filename = fd.askopenfilename(
            title="Открыть файл", initialdir=initialdir, filetypes=filetypes
        )
        if filename:
            # Если выбран верный файл, то окрашиваем в зеленый и удаляем ошибку
            if filename.endswith(".xlsx"):
                btn_file["bg"] = "LimeGreen"
                try:
                    self.text_error.destroy()
                except AttributeError:
                    pass
                try:
                    # Если была ошибка при открытии файла эксель
                    self.troubles_excel.destroy()
                except AttributeError:
                    pass
                try:
                    self.file_processing.destroy()
                except AttributeError:
                    pass
                try:
                    self.error_panel.destroy()
                except AttributeError:
                    pass
                btn_file["text"] = "Файл выбран"
                self.update()
                file.path = filename
            else:
                try:
                    # Если была ошибка при открытии файла эксель
                    self.troubles_excel.destroy()
                except AttributeError:
                    pass
                try:
                    self.file_processing.destroy()
                except AttributeError:
                    pass
                try:
                    self.error_panel.destroy()
                except AttributeError:
                    pass
                # Кнопка окрасится в красный, если не верный формат
                btn_file["bg"] = "Crimson"
                self.text_error = tk.Label(
                    self,
                    font=("Microsoft Sans Serif", 16),
                    text='Выберите файл в формате ".xlsx". Возможно, вы используете устаревший\n формат.',
                    fg="#FF1919",
                )

                self.text_error.place(x=10, y=110)


if __name__ == "__main__":
    app = App()
    Emk(app)
    Bunk(app)
    Phone(app)
    Operation(app)
    Service(app)
    import pyi_splash
    pyi_splash.close()
    app.mainloop()
