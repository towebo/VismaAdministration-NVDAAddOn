﻿import os
import api
import wx
from ctypes import *
import config
import gui
import appModuleHandler
from NVDAObjects.UIA import UIA
from NVDAObjects.window import Window
import UIAHandler
import tones
import winUser
import winKernel
import watchdog
from NVDAObjects.IAccessible import IAccessible, ContentGenericClient
import oleacc
import speech

import controlTypes
import ui
import tones
from logHandler import log



# Constants
MODULE_OFFER = "Offert"
MODULE_ORDER = "Order"
MODULE_INVOICE = "Kundfaktura"
MODULE_AGREEMENT_TEMPLATE = "Avtalsmall"
MODULE_AGREEMENT = "Avtal"
MODULE_ARTICLE = "Artiklar och tjänster"
MODULE_CUSTOMER = "Kunder"
MODULE_SUPPLIER = "Leverantörer"
MODULE_GENERAL = "Allmän"
MODULE_BOOKINGS = "Beställning"
MODULE_DELIVERYNOTE = "Inkommande följesedel"
MODULE_SUPPLIERINVOICE = "Leverantörsfaktura"
MODULE_CONTACTS = "Kontakter"
MODULE_PRICELIST = "Prislistor"
MODULE_MANUALDELIVERYIN = "Manuell inleverans"
MODULE_MANUALDELIVERYOUT = "Manuell utleverans"
MODULE_INVENTORY = "Inventering"
MODULE_PRICEIMPORT = "Prisinläsning"
MODULE_PRICERECALCULATION_ESTIMATEDPRICES = "Prisomräkning kalkylpriser"
MODULE_PRICERECALCULATION_SUPPLIERPRICES = "Prisomräkning leverantörspriser"
MODULE_INGOINGBALLANCE = "Ingående balans"
MODULE_VERIFICATIONS = "Verifikationer"

MODULE_RESULTBUDGET= "Resultatbudget"
MODULE_RESULTPROGNOZE = "Resultatprognos"
MODULE_BALLANCEBUDGET = "Balansbudget"
MODULE_BALLANCEPROGNOZE = "Balansprognos"
MODULE_ACCOUNTPLAN = "Kontoplan"
MODULE_DISCOUNTAGREEMENTS = "Kundrabatter"
MODULE_MEMBERS = "Medlemmar"

# https://www.pinvoke.net/default.aspx/Constants.TCM_
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

fn = os.path.dirname(os.path.abspath(__file__)) + "\\..\\..\\Data\\\\controls.txt"
with open(fn, 'r') as f:
    ctrllines = f.read().splitlines()
    

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


class AppModule(appModuleHandler.AppModule):

    last_module = None

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
        last_module = self.getCurrentVismaModule(obj)
        txt = self.getControlName(obj)

        #ui.message("%d, %s" % (last_module, obj.windowControlID))


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
                clsList.insert(0, AdminSystemDialog)
            elif isinstance(obj, IAccessible) and obj.IAccessibleRole == oleacc.ROLE_SYSTEM_PAGETABLIST:
                clsList.insert(0, AdminTabControl)
            elif isinstance(obj, IAccessible) and obj.windowClassName.startswith("BCGPControlBar") and obj.IAccessibleRole == oleacc.ROLE_SYSTEM_CLIENT:
                clsList.insert(0, AdminControlBar)
        except Exception as e:
            ui.message("Fel: %s" % e)
            

            
        
    def script_readControlInfo(self, obj):
        wnd = api.getFocusObject()
        module = self.getCurrentVismaModule(wnd)
        
        #wnd = wnd.parent.parent.parent.parent.parent.parent.parent.parent.parent
        
        txt = "Name: %s, WindowControlID: %d, WindowClassName: %s, VismaModule: %s" % ( wnd.name, wnd.windowControlID, wnd.windowClassName, module )
        ui.message(txt)

    def getCurrentVismaModule(self, ctrl):
        try:
            wnd = ctrl
            if ctrl == None:
                wnd = api.getFocusObject()
            if wnd is None:
                return self.appName
            while wnd is not None and isinstance(wnd, Window) and not wnd.windowClassName.startswith('Afx:'):
                wnd = wnd.parent
                if wnd is None:
                    return self.appName
                if wnd.windowClassName == "#32770" and wnd.windowText == "Massuppdatering":
                    return self.last_module
                #log.debug("%s, %s, %d" % ( wnd.windowClassName, wnd.windowText, wnd.windowControlID ))
                #ui.message(wnd.windowClassName)
            module = None
            wndtxt = wnd.windowText.lower()
            if wndtxt.startswith('order'):
                module = MODULE_ORDER
            elif wndtxt.startswith('kundfakt'):
                module = MODULE_INVOICE
            elif wndtxt.startswith('offerter'):
                module = MODULE_OFFER
            elif wndtxt.startswith('kunder'):
                module = MODULE_CUSTOMER
            elif wndtxt.startswith('leverantörer'):
                module = MODULE_SUPPLIER
            elif wndtxt.startswith('artiklar'):
                module = MODULE_ARTICLE
            elif wndtxt.startswith('beställningar'):
                module = MODULE_BOOKINGS
            elif wndtxt.startswith('inkommande följ'):
                module = MODULE_DELIVERYNOTE
            elif wndtxt.startswith('leverantörsfakturor'):
                module = MODULE_SUPPLIERINVOICE
            elif wndtxt.startswith('kontakter'):
                module = MODULE_CONTACTS
            elif wndtxt.startswith('avtalsmall'):
                module = MODULE_AGREEMENT_TEMPLATE
            elif wndtxt.startswith('avtal'):
                module = MODULE_AGREEMENT
            elif wndtxt.startswith('försäljningsprislis'):
                module = MODULE_PRICELIST
            elif wndtxt.startswith('manuella inleveranser'):
                module = MODULE_MANUALDELIVERYIN
            elif wndtxt.startswith('manuella utleveranser'):
                module = MODULE_MANUALDELIVERYOUT
            elif wndtxt.startswith('inventering'):
                module = MODULE_INVENTORY
            elif wndtxt.startswith('prisinläsning'):
                module = MODULE_PRICEIMPORT
            elif wndtxt.startswith('prisomräkning kalkyl'):
                module = MODULE_PRICERECALCULATION_ESTIMATEDPRICES
            elif wndtxt.startswith('prisomräkning lev'):
                module = MODULE_PRICERECALCULATION_SUPPLIERPRICES
            elif wndtxt.startswith('ingående balans'):
                module = MODULE_INGOINGBALLANCE
            elif wndtxt.startswith('verifikationer'):
                module = MODULE_VERIFICATIONS
            elif wndtxt.startswith('resultatbudg'):
                module = MODULE_RESULTBUDGET
            elif wndtxt.startswith('resultatprognos'):
                module = MODULE_RESULTPROGNOZE
            elif wndtxt.startswith('balansbudget'):
                module = MODULE_BALLANCEBUDGET
            elif wndtxt.startswith('balansprognos'):
                module = MODULE_BALLANCEPROGNOZE
            elif wndtxt.startswith('kontoplan'):
                module = MODULE_ACCOUNTPLAN
            elif wndtxt.startswith('kundrabatter'):
                module = MODULE_DISCOUNTAGREEMENTS
            elif wndtxt.startswith('medlemmar'):
                module = MODULE_MEMBERS
            else:
                module = "Fönstertext " + wndtxt

            self.last_module = module
            return module
        except Exception as e:
            ui.message("Fel: %s" % e)


    def getControlName(self, obj):
        try:
            global ctrllines
            wndclasslines = [k for k in ctrllines if obj.windowClassName in k]
            if len(wndclasslines) == 0:
                return None
            idlines = [k for k in wndclasslines if ("%d" % obj.windowControlID) in k]
            if len(idlines) == 0:
                return None
            elif len(idlines) >= 2:
                module = self.getCurrentVismaModule(obj)
                #module = self.last_module
                idlines = [k for k in idlines if module in k]
                if len(idlines) == 0:
                    return None
            lineparts = idlines[0].split('\t')
            return lineparts[3]
        except Exception as e:
            ui.message("Fel: %s" % e)


    def doReadVismaCommands(self):
        try:
            modul = self.getCurrentVismaModule(None)
            ui.message(modul)
            fn = os.path.dirname(os.path.abspath(__file__)) + "\\..\\..\\Data\\shortcuts.txt"
            with open(fn, 'r') as f:
                helplines = f.read().splitlines()
                modulehelplines = [k for k in helplines if modul in k]
                ui.message("Det finns %d kortkommandon" % len(modulehelplines))
                for line in modulehelplines:
                    lineparts = line.split('\t')
                    keys = ""
                    for c in lineparts[3]:
                        keys = keys + c + " "
                    txt = "%s %s, Alt + K %s, %s" % (lineparts[1], lineparts[2], keys, lineparts[4])
                    ui.message(txt)

                modul = "Allmänt"
                generalhelplines = [k for k in helplines if modul in k]
                ui.message("Det finns %d generella kortkommandon" % len(generalhelplines))
                for line in generalhelplines:
                    lineparts = line.split('\t')
                    keys = ""
                    for c in lineparts[3]:
                        keys = keys + c + " "
                    txt = "%s %s, Alt + K %s, %s" % (lineparts[1], lineparts[2], keys, lineparts[4])
                    ui.message(txt)

        except Exception as e:
            ui.message("Fel: %s" % e)


    def script_readVismaCommands(self, gesture):
        self.doReadVismaCommands()

    __gestures = {
        "kb:NVDA+h": "readVismaCommands",
        "kb:NVDA+i": "readControlInfo"
    }


class VismaSafGrid(UIA):


    def event_gainFocus(self):
        try:
            if self.name != "Grid":
                ui.message(self.name)
            if config.conf['VismaAdministration']['sayNumGridRows']:
                ui.message("Listan har %d rader" % self._get_rowCount())
            #ui.message("Rad %d är markerad" % self._get_rowNumber())
            self.ReadGridSelection()
        except Exception as e:
            ui.message("Fel: %s" % e)
    
    def event_UIA_selectionInvalidated(self):
        self.ReadGridSelection()
        
    def event_UIA_AutomationFocusChanged(self, obj, nextHandler):
        self.ReadGridSelection()
        nextHandler

    def script_readNumGridRows(self, gesture):
        # Pass the keystroke along
        #gesture.send()
        ui.message("Listan har %d rader" % self._get_rowCount())
        ui.message("Rad %d är markerad" % self._get_rowNumber())

    def script_readGridSelection(self, gesture):
        # Pass the keystroke along
        #gesture.send()
        self.ReadGridSelection()
        
        
    def ReadGridSelection(self):
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
                    if self.windowControlID == 7319 or self.windowControlID == 7320:
                        if coltxt == "Kol 0":
                            coltxt = ""
                    # The list of companies in a backup
                    elif self.windowControlID == 7221:
                        if coltxt == "Kol 0":
                            coltxt = "Markerad"
                            ischeckbox = True
                    # Checkboxes
                    if coltxt == "Markering" or coltxt == "Inaktiv" or coltxt == "Aktivt":
                        ischeckbox = True
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
                        if ischeckbox:
                            if valtxt == "1":
                                valtxt = "Ja"
                                valtxt = ""
                            elif valtxt == "0":
                                valtxt = "Nej"
                                valtxt = ""
                                coltxt = ""
                    ui.message("%s, %s" % (coltxt, valtxt))
                    
            # A grid that selects each cell
            else:
                cursel = selpat.GetCurrentSelection()
                if cursel.Length > 0:
                    selement = cursel.GetElement(0)
                    coltxt = selement.CurrentName
                    ischeckbox = False
                    # Checkboxes
                    if coltxt == "Skriv" or coltxt == "Skriv order" or coltxt == "Skriv följ" or coltxt == "Restn ej":
                        ischeckbox = True
                    
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
            ui.message("Fel: %s"%e)
            
    
    __gestures = {
        "kb:NVDA+m": "readGridSelection",
        "kb:NVDA+r": "readNumGridRows"
        

    }


class VismaAdministrationSettingsPanel(gui.SettingsPanel):
    # Translators: the label/title for the Visma Administration settings panel.
    title = _('Visma Administration')

    def makeSettings(self, settingsSizer):
        helper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
        # Translators: the label for the preference to choose the maximum number of stored history entries
        #maxHistoryLengthLabelText = _('&Maximum number of history entries (requires NVDA restart to take effect)')
        #self.maxHistoryLengthEdit = helper.addLabeledControl(maxHistoryLengthLabelText, nvdaControls.SelectOnFocusSpinCtrl, min=1, max=5000, initial=config.conf['speechHistory']['maxHistoryLength'])
        # Translators: the label for the preference to trim whitespace from the start of text
        self.debugModeCB = helper.addItem(wx.CheckBox(self, label=_('Utvecklarläge (Ger onödigt mycket information)')))
        self.debugModeCB.SetValue(config.conf['VismaAdministration']['debugMode'])
        # Translators: Nothig to see here
        self.sayNumGridRowsCB = helper.addItem(wx.CheckBox(self, label=_('Säg antalet rader för listor automatiskt')))
        self.sayNumGridRowsCB.SetValue(config.conf['VismaAdministration']['sayNumGridRows'])

    def onSave(self):
        config.conf['VismaAdministration']['debugMode'] = self.debugModeCB.GetValue()
        config.conf['VismaAdministration']['sayNumGridRows'] = self.sayNumGridRowsCB.GetValue()

class AdminTabControl(IAccessible):

    def initOverlayClass(self):
        speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)

    def event_gainFocus(self):
        try:
            # Ensure the selected tab text is spoken instead of 'Fliksidor'
            self.name = self.get_tab_text(-1)
            speech.speakObject(self, reason=controlTypes.OutputReason.FOCUS)
        except Exception as e:
            ui.message("Fel: %s" % e)
        
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

class AdminSystemDialog(IAccessible):

    def _get_name(self):
        return None
    
    def _get_value(self):
        return None

    def _get_roleText(self):
        return " "

    def _get_description(self):
        # Skip the announcement of all controls on the dialog page
        return None 


class AdminControlBar(IAccessible):

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
            ui.message("Fel: %s" % e)
    
    def _get_name(self):
        return None

    def _get_description(self):
        return None

    