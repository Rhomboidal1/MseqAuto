from pywinauto.application import Application

app = Application(backend='uia').start('notepad.exe').connect(title='Untitled - Notepad',timeout=100)
#app = Application(backend='uia').connect(title='Untitled - Notepad',timeout=100)
#app.UntitledNotepad.print_control_identifiers()
app.UntitledNotepad.child_window(title="Text Editor", auto_id="15", control_type="Edit").wrapper_object()