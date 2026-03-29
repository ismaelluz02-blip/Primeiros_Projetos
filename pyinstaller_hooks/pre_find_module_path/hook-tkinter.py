# Hook local para nao excluir tkinter mesmo com Tcl/Tk do host inconsistente.
# O pacote tkinter sera coletado manualmente pelo Analysis (hiddenimports + binaries/datas).
def pre_find_module_path(api):
    return
