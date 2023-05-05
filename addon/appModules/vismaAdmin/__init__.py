import os
import api
import wx
from ctypes import *
import config
import gui
import appModuleHandler
import addonHandler
import locale
from NVDAObjects.UIA import UIA
from NVDAObjects.window import Window
import UIAHandler
import winKernel
import watchdog
from NVDAObjects.IAccessible import IAccessible, ContentGenericClient
import oleacc
import speech
import scriptHandler
from scriptHandler import script

import controlTypes
import ui
import tones
from logHandler import log


 #https://www.pinvoke.net/default.aspx/Constants.TCM_
TCM_FIRST = 0x1300
TCM_GETITEMA = TCM_FIRST + 5
TCM_GETITEMW = TCM_FIRST + 60
TCM_GETCURSEL = TCM_FIRST + 11

TCIF_TEXT = 0x0001
TCIF_IMAGE = 0x0002
TCIF_RTLREADING = 0x0004
TCIF_PARAM = 0x0008
TCIF_STATE = 0x0010


# For convenience.
ExtraUIAEvents = {
#    UIAHandler.UIA_AutomationFocusChangedEventId: "UIA_AutomationFocusChanged",
    UIAHandler.UIA_Selection_InvalidatedEventId: "UIA_selectionInvalidated"

}

ctrllines = []
module_lines = []

fn = os.path.dirname(os.path.abspath(__file__)) + "\\..\\..\\data\\\\controls.txt"
with open(fn, 'r') as f:
    ctrllines = f.read().splitlines()

fn = os.path.dirname(os.path.abspath(__file__)) + "\\..\\..\\data\\\\modules.txt"
with open(fn, 'r') as f:
    module_lines = f.read().splitlines()
    

last_tab_text = ""

class TCITEMWStruct(Structure):
    _fields_=[
        ("mask", wintypes.UINT),
        ('state', wintypes.DWORD),
        ('stateMask', wintypes.DWORD),
        ('text', wintypes.LPWSTR),
        ('textMax', wintypes.INT),
        ('image', wintypes.INT),
        ('lParam', wintypes.LPARAM),
    ]

addonHandler.initTranslation()

# for debug logging
DEBUG = False

def debugLog(message):
	if DEBUG:
		log.info(message)



class AppModule(appModuleHandler.AppModule):

    last_module = ""

    def __init__(self, *args, **kwargs):
        super(AppModule, self).__init__(*args, **kwargs)
        # Add a series of events instead of doing it one at a time.
        # Some events are only available in a specific build range and/or while a specific version of IUIAutomation interface is in use.
        for event, name in ExtraUIAEvents.items():
            if event not in UIAHandler.UIAEventIdsToNVDAEventNames:
                UIAHandler.UIAEventIdsToNVDAEventNames[event] = name
                UIAHandler.handler.clientObject.addAutomationEventHandler(event,UIAHandler.handler.rootElement,UIAHandler.TreeScope_Subtree,UIAHandler.handler.baseCacheRequest,UIAHandler.handler)
                
        confspec = {
            'debugMode': 'boolean(default=false)',
            'sayNumGridRows': 'boolean(default=false)'
        }
        config.conf.spec['VismaAdministration'] = confspec
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(VismaAdministrationSettingsPanel)


    def event_NVDAObject_init(self, obj):
        if not isinstance(obj, Window):
            pass
        module = self.getCurrentVismaModule(obj)
        # Only store the module if it's something else than the app so it's useful.
        if module != obj.appModule.appName:
            self.last_module = module
        txt = self.getControlName(obj, self.last_module)
        if txt is None or txt =="":
            txt = self.getControlName(obj, module)

        if txt != None:
            obj.name = txt
        else:
            if config.conf['VismaAdministration']['debugMode']:
                if obj.windowClassName == "Button" or obj.windowClassName == "Edit" or obj.windowClassName == "ComboBox":
                    try:
                        obj.name = "%s, %s, %d, %s" % ( obj.name, obj.displayText, obj.windowControlID, obj.windowClassName )
                    except:
                        obj.name = "%d, %s" % ( obj.windowControlID, obj.windowClassName )

            
    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        try:
            if obj.windowClassName == "SafGrid" or obj.windowClassName == "AfxWnd140s":
                clsList.insert(0, VismaSafGrid)
            elif isinstance(obj, IAccessible) and obj.IAccessibleRole == oleacc.ROLE_SYSTEM_DIALOG:
                clsList.insert(0, VismaSystemDialog)
            elif isinstance(obj, IAccessible) and obj.IAccessibleRole == oleacc.ROLE_SYSTEM_PAGETABLIST:
                clsList.insert(0, SysTabControl32)
            elif isinstance(obj, IAccessible) and obj.windowClassName.startswith("BCGPControlBar") and obj.IAccessibleRole == oleacc.ROLE_SYSTEM_CLIENT:
                clsList.insert(0, VismaControlBar)
            elif isinstance(obj, IAccessible) and obj.IAccessibleRole == oleacc.ROLE_SYSTEM_CHECKBUTTON:
                clsList.insert(0, SystemCheckButton)
        except Exception as e:
            log.info("Fel i chooseNVDAObjectOverlayClasses: %s" % e)
            ui.message("Fel i chooseNVDAObjectOverlayClasses: %s" % e)
    
    @script(
        # Translators: Gesture description
        description=_("Says info about the current control. Press twice to copy the information to the clipboard."),
        category=_("Visma Administration"), 
        gesture="kb:NVDA+i"
    )
    def script_readControlInfo(self, obj):
        wnd = api.getFocusObject()
        module = self.getCurrentVismaModule(wnd)
        
        txt = "%s\t%d\t%s\t%s\t" % ( wnd.windowClassName, wnd.windowControlID, module,wnd.name )

        isSameScript = scriptHandler.getLastScriptRepeatCount()
        if isSameScript  == 0:
            ui.message(txt)
        else:
            if api.copyToClip(txt):
                ui.message("%s kopierat till klippbordet" % txt)

    def getCurrentVismaModule(self, ctrl):
        try:
            wnd = ctrl
            if ctrl == None:
                wnd = api.getFocusObject()
            if wnd is None:
                return self.appName
            wndtxt = wnd.windowText
            while wnd is not None and isinstance(wnd, Window) and not wnd.windowClassName.startswith('Afx'):
                try:
                    wnd = wnd.parent
                    if wnd is not None and wnd.windowText != "":
                        wndtxt = wnd.windowText
                    if wnd is not None and wnd.windowClassName == "#32770" and wnd.windowText == "Massuppdatering":
                        return self.last_module
                except Exception as e:
                    pass
                    #log.info("Fel i getCurrentVismaModule när wnd (%s, %d, %s) parent skulle hämtas: %s" % (wnd.windowClassName, wnd.windowControlID, wnd.name, e))
                #if wnd is None:
                #    try:
                #        return self.appName
                #    except Exception as e:
                #        #log.info("Fel i getCurrentVismaModule när self.appName skulle hämtas: %s" % e)
                #        #ui.message("Fel inträffade")
                #        return None
                
            module = wndtxt
            wndtxt = wndtxt.lower()
            
            global module_lines
            lines = [k for k in module_lines if ("%s\t" % wndtxt) in k]
            if len(lines ) == 0:
                log.info("Ny modul: %s" % module)
                return module
            lineparts = lines[0].split('\t')
            module = lineparts[1]
            self.last_module = module
            return module
        except Exception as e:
            log.info("Fel i getCurrentVismaModule: %s" % e)
            ui.message("Fel i getCurrentVismaModule: %s" % e)


    def getControlName(self, obj, module):
        try:
            global ctrllines
            wndclasslines = [k for k in ctrllines if obj.windowClassName in k]
            if len(wndclasslines) == 0:
                return None
            idlines = [k for k in wndclasslines if ("\t%d\t" % obj.windowControlID) in k]
            if len(idlines) == 0:
                return None
            elif len(idlines) == 1:
                lineparts = idlines[0].split('\t')
                ctrl_module = lineparts[2]
                if ctrl_module == None or ctrl_module == "":
                    return lineparts[3]
                if ctrl_module != module:
                    return None
            else:
                idlines = [k for k in idlines if module in k]
                if len(idlines) == 0:
                    return None
            lineparts = idlines[0].split('\t')
            return lineparts[3]
        except Exception as e:
            log.info("Fel i getControlName: %s" % e)
            ui.message("Fel i getControlName: %s" % e)


    def doReadVismaCommands(self):
        try:
            modul = self.getCurrentVismaModule(None)
            ui.message(modul)
            fn = os.path.dirname(os.path.abspath(__file__)) + "\\..\\..\\data\\shortcuts.txt"
            with open(fn, 'r') as f:
                helplines = f.read().splitlines()
                
                isSameScript = scriptHandler.getLastScriptRepeatCount()
                if isSameScript  == 0:
                    modulehelplines = [k for k in helplines if modul in k]
                    ui.message(_("There are %d keyboard shortcuts") % len(modulehelplines))
                    for line in modulehelplines:
                        lineparts = line.split('\t')
                        keys = ""
                        for c in lineparts[3]:
                            keys = keys + c + " "
                        txt = "%s %s, Alt + K %s, %s" % (lineparts[1], lineparts[2], keys, lineparts[4])
                        ui.message(txt)
                elif isSameScript  == 1:
                    modul = "Allmänt"
                    generalhelplines = [k for k in helplines if modul in k]
                    ui.message(_("There are %d common keyboard shortcuts") % len(generalhelplines))
                    for line in generalhelplines:
                        lineparts = line.split('\t')
                        keys = ""
                        for c in lineparts[3]:
                            keys = keys + c + " "
                        txt = "%s %s, Alt + K %s, %s" % (lineparts[1], lineparts[2], keys, lineparts[4])
                        ui.message(txt)
        except Exception as e:
            log.info("Fel i doReadVismaCommands: %s" % e)
            ui.message("Fel i doReadVismaCommands: %s" % e)


    @script(
        # Translators: Gesture description
        description=_("Announces keyboard shortcuts for current view"),
        category=_("Visma Administration"), 
        gesture="kb:NVDA+h"
    )
    def script_readVismaCommands(self, gesture):
        self.doReadVismaCommands()

class VismaSafGrid(UIA):

    __changeItemGestures = (
        "kb:space",
        )
    
    def initOverlayClass(self):
        for gesture in self.__changeItemGestures:
            self.bindGesture(gesture, "changeItem")


    def script_changeItem(self,gesture):
        gesture.send()
        self.ReadGridSelection()


    def event_gainFocus(self):
        try:
            if self.name != "Grid":
                ui.message(self.name)
            if config.conf['VismaAdministration']['sayNumGridRows']:
                ui.message("Listan har %d rader" % self._get_rowCount())
            #ui.message("Rad %d är markerad" % self._get_rowNumber())
            self.ReadGridSelection()
        except Exception as e:
            log.info("Fel i VismaSafGrid.gainFocus: %s" % e)
            ui.message("Fel i VismaSafGrid.gainFocus: %s" % e)
    
    def event_UIA_selectionInvalidated(self):
        try:
            self.ReadGridSelection()
        except Exception as e:
            log.info("Fel i VismaSafGrid.event_UIA_selectionInvalidated: %s" % e)
            ui.message("Fel i VismaSafGrid.event_UIA_selectionInvalidated: %s" % e)
        
    def event_UIA_AutomationFocusChanged(self, obj, nextHandler):
        try:
            self.ReadGridSelection()
        except Exception as e:
            log.info("Fel i VismaSafGrid.event_UIA_AutomationFocusChanged: %s" % e)
            ui.message("Fel i VismaSafGrid.event_UIA_AutomationFocusChanged: %s" % e)
        nextHandler

    @script(
        # Translators: Gesture description
        description=_("Says the number of rows in the current list"),
        category=_("Visma Administration"), 
        gesture="kb:NVDA+r"
    )
    def script_readNumGridRows(self, gesture):
        # Pass the keystroke along
        #gesture.send()
        ui.message("Listan har %d rader" % self._get_rowCount())
        #ui.message("Rad %d är markerad" % self._get_rowNumber())

    @script(
        # Translators: Gesture description
        description=_("Reads the selection of the list"),
        category=_("Visma Administration"), 
        gesture="kb:NVDA+m"
    )
    def script_readGridSelection(self, gesture):
        # Pass the keystroke along
        #gesture.send()
        self.ReadGridSelection()
        
        
    def ReadGridSelection(self):
        checkbox_cols = ["Markering", "Inaktiv", "Aktivt", "Makulerad", "Fakturerad", "Skriv", "Skriv order", "Skriv följ", "Restn ej", "Skickad", "Levererad", "Order"]
        
        try:
            #gridpat = nav._getUIAPattern(UIAHandler.UIA_GridPatternId,UIAHandler.IUIAutomationGridPattern)
            #tablepat = nav._getUIAPattern(UIAHandler.UIA_TablePatternId,UIAHandler.IUIAutomationTablePattern)
            #tableitempat = nav._getUIAPattern(UIAHandler.UIA_TableItemPatternId,UIAHandler.IUIAutomationTableItemPattern)
            #selpat = nav._getUIAPattern(UIAHandler.UIA_SelectionPatternId,UIAHandler.IUIAutomationSelectionPattern)
            selpat = self._getUIAPattern(UIAHandler.UIA_SelectionPatternId,UIAHandler.IUIAutomationSelectionPattern)
            
            # Is it a grid with row selection?
            if selpat.CurrentCanSelectMultiple:
                cursel = selpat.GetCurrentSelection()
                for i in range(cursel.Length):
                    selement = cursel.GetElement(i)
                    coltxt = selement.CurrentName
                    ischeckbox = False
                    # Company list
                    if self.windowControlID == 7407:
                        if coltxt == "Kol 0":
                            coltxt = U"Företag"
                        elif coltxt == U"Företagsnamn":
                            coltxt = ""
                        elif coltxt == "Kol 1":
                            coltxt = "Mapp"
                    # Aktiva kolumner i kolumnredigeringsdialogen
                    elif self.windowControlID == 7319 or self.windowControlID == 7320:
                        if coltxt == "Kol 0":
                            coltxt = ""
                    # The list of companies in a backup
                    elif self.windowControlID == 7221:
                        if coltxt == "Kol 0":
                            coltxt = "Markerad"
                            ischeckbox = True
                    # Checkboxes
                    ischeckbox = checkbox_cols.count(coltxt) > 0
                    if ischeckbox:
                        if coltxt == "Markering":
                            coltxt = "Markerad"
                    punk=selement.getCurrentPattern(UIAHandler.UIA_ValuePatternId)
                    if punk:
                        valpat =punk.QueryInterface(UIAHandler.IUIAutomationValuePattern)
                        if valpat:
                            try:
                                valtxt = valpat.CurrentValue;
                            except:
                                valtxt = ""

                        doread = True
                        if ischeckbox:
                            if valtxt == "1":
                                valtxt = "Ja"
                            elif valtxt == "0":
                                valtxt = "Nej"
                                if coltxt == "Markerad" or coltxt == "Inaktiv":
                                    doread = False
                        if doread:
                            ui.message("%s, %s" % (coltxt, valtxt))
                    
            # A grid that selects each cell
            else:
                cursel = selpat.GetCurrentSelection()
                if cursel.Length > 0:
                    selement = cursel.GetElement(0)
                    coltxt = selement.CurrentName
                    ischeckbox = False
                    # Checkboxes
                    ischeckbox = checkbox_cols.count(coltxt) > 0
                    punk = selement .getCurrentPattern(UIAHandler.UIA_ValuePatternId)
                    if punk:
                        valpat =punk.QueryInterface(UIAHandler.IUIAutomationValuePattern)
                        try:
                            valtxt = valpat.CurrentValue;
                        except:
                            valtxt = ""

                        if ischeckbox:
                            if valtxt == "1":
                                valtxt = "Ja"
                            elif valtxt == "0":
                                valtxt = "Nej"

                    ui.message("%s, %s" % (coltxt, valtxt))
        except Exception as e:
            log.info("Fel i readGridSelection: %s"%e)
            ui.message("Fel i readGridSelection: %s"%e)
            



class VismaAdministrationSettingsPanel(gui.SettingsPanel):
    # Translators: the label/title for the Visma Administration settings panel.
    title = _('Visma Administration')

    def makeSettings(self, settingsSizer):
        helper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
        # Translators: Label for checkbox that controls if running in developer mode
        self.debugModeCB = helper.addItem(wx.CheckBox(self, label=_('Developer Mode')))
        self.debugModeCB.SetValue(config.conf['VismaAdministration']['debugMode'])
        # Translators: Checkbox in the settings panel
        self.sayNumGridRowsCB = helper.addItem(wx.CheckBox(self, label=_('Say number of rows in lists automatically')))
        self.sayNumGridRowsCB.SetValue(config.conf['VismaAdministration']['sayNumGridRows'])

    def onSave(self):
        config.conf['VismaAdministration']['debugMode'] = self.debugModeCB.GetValue()
        config.conf['VismaAdministration']['sayNumGridRows'] = self.sayNumGridRowsCB.GetValue()

class SysTabControl32(IAccessible):

    def initOverlayClass(self):
        global last_tab_text
        txt = self.name
        if txt != last_tab_text:
            last_tab_text = txt
            speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)

    def event_gainFocus(self):
        try:
            # Ensure the selected tab text is spoken instead of 'Fliksidor'
            self.name = self.get_tab_text(-1)
            speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)
        except Exception as e:
            log.info("Fel i VismaTabControl.gainFocus: %s" % e)
            ui.message("Fel i VismaTabControl.gainFocus: %s" % e)

    def _get_name(self):
        return self.get_tab_text(-1)
    
    def _get_value(self):
        return None

    def _get_roleText(self):
        return " " # Skip the announcement of 'Flikar'

    def get_tab_text(self,idx):
        tabidx = idx
        if idx == -1:
            tabidx = watchdog.cancellableSendMessage(self.windowHandle, TCM_GETCURSEL, 0, 0)
        bufLen = 256
        info = TCITEMWStruct()
        info.mask = TCIF_TEXT
        info.textMax = bufLen - 1
        
        processHandle = self.processHandle
        internalBuf = winKernel.virtualAllocEx(processHandle, None, bufLen, winKernel.MEM_COMMIT, winKernel.PAGE_READWRITE)
        try:
            info.text = internalBuf
            internalInfo = winKernel.virtualAllocEx(processHandle, None, sizeof(info), winKernel.MEM_COMMIT, winKernel.PAGE_READWRITE)
            try:
                winKernel.writeProcessMemory(processHandle, internalInfo, byref(info), sizeof(info), None)
                got_it = watchdog.cancellableSendMessage(self.windowHandle, TCM_GETITEMW, tabidx, internalInfo)
            finally:
                winKernel.virtualFreeEx(processHandle, internalInfo, 0, winKernel.MEM_RELEASE)
            buf = create_unicode_buffer(bufLen)
            winKernel.readProcessMemory(processHandle, internalBuf, buf, bufLen, None)
        finally:
            winKernel.virtualFreeEx(processHandle, internalBuf, 0, winKernel.MEM_RELEASE)
        if got_it:
            text = buf.value
            return text
        return None            

class VismaSystemDialog(IAccessible):

    def _get_name(self):
        return None
    
    def _get_value(self):
        return None

    def _get_roleText(self):
        return " "

    def _get_description(self):
        # Skip the announcement of all controls on the dialog page
        return None 


class VismaControlBar(IAccessible):

    def initOverlayClass(self):
        try:
            speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)
            if self.value is None:
                return
            btn = self.simpleFirstChild
            if btn is None:
                return
            while btn is not None and btn.IAccessibleRole != oleacc.ROLE_SYSTEM_PUSHBUTTON:
                btn = btn.simpleNext
            if btn is not None:
                api.moveMouseToNVDAObject(btn)
        except Exception as e:
            log.info("Fel i AdminControlBar.initOverlayClass: %s" % e)
            ui.message("Fel i AdminControlBar.initOverlayClass: %s" % e)
    
    def _get_name(self):
        return None

    def _get_description(self):
        return None

class SystemCheckButton(IAccessible):

    __changeItemGestures = (
        "kb:space",
        )
    
    def initOverlayClass(self):
        for gesture in self.__changeItemGestures:
            self.bindGesture(gesture, "changeItem")


    def script_changeItem(self,gesture):
        gesture.send()
        speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)

