import flet as ft
from exlWrapper import ExcelWrapper

exl = ExcelWrapper()
def main(page: ft.Page):
    pathFile = ''

    def on_dialog_result_save(e: ft.FilePickerResultEvent):
        global pathFile
        pathFile = e.path
        exl.save(pathFile+'.xlsx')
        page.update()

    def on_dialog_result(e: ft.FilePickerResultEvent):
        global pathFile
        pathFile = e.files[0].path
        path.value = pathFile
        exl.feed(['Вложения', 'Последний раз обновлено', 'Статус', 'Наименование сервисного центра'], ['ПЭ: дата время', 'ПЭ: Комментарий', 'ПЭ: наработка м/ч'], pathFile)
        path.update
        exl.format()
        dlg_waiting.open = False
        page.dialog = dlg_modal
        dlg_modal.open = True
        page.update()

    def close_dlg(e): # НЕМНОГО КОСТЫЛЕЙ
        dlg_modal.open = False
        dlg_modal_format.open = False
        page.update()

    def on_hover(e):
        e.control.bgcolor = primary_color if e.data == "true" else accent_color
        e.control.color = "black" if e.data == "true" else primary_color        
        e.control.update()

    def on_save(e):
        file_picker_save.save_file()
        page.update()

    def on_file_pick(e):
        file_picker.pick_files(allow_multiple=False)
        page.dialog = dlg_waiting
        dlg_waiting.open = True
        page.update()

    # Цветовая палитра
    accent_color = "#ce2a2e"
    primary_color = "#ffffff"

    page.title = "Тестовая версия ПТЗ"


    file_picker = ft.FilePicker(on_result=on_dialog_result)
    file_picker_save = ft.FilePicker(on_result=on_dialog_result_save)

    dlg_waiting = ft.AlertDialog( # Уведомление о процессе подгрузки файла
        modal=True,
        title=ft.Text("Файл подгружается"),
        actions_alignment=ft.MainAxisAlignment.END,)

    dlg_modal = ft.AlertDialog( # Уведомление о завершении подгрузки файла
        modal=True,
        title=ft.Text("Отчет выбран"),
        actions=[
            ft.TextButton("ОК", on_click=close_dlg),
        ],
        actions_alignment=ft.MainAxisAlignment.END,)

    dlg_modal_format = ft.AlertDialog( # Уведомление о завершении форматирования файла
        modal=True,
        title=ft.Text("Отчет отформатирован"),
        actions=[
            ft.TextButton("ОК", on_click=close_dlg),
        ],
        actions_alignment=ft.MainAxisAlignment.END,)

    b1 = ft.ElevatedButton("Выбрать файл", bgcolor = accent_color, color =primary_color,
                           on_click=on_file_pick,
                           on_hover=on_hover,icon = "ATTACH_FILE")
    path = ft.Text("Здесь появится путь к файлу",
                   text_align = ft.TextAlign.CENTER)
    
    b3 = ft.ElevatedButton(text="Сохранить файл", bgcolor = accent_color,
                           color = primary_color, on_hover=on_hover,
                           icon = "SAVE", on_click=on_save)
    
    page.overlay.append(file_picker)
    page.overlay.append(file_picker_save)
    page.update()
    page.add(b1, path, b3)
ft.app(target=main)

