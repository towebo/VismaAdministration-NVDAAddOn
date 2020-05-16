import os
import api
import wx
import config
import gui
import appModuleHandler
from NVDAObjects.UIA import UIA
from NVDAObjects.window import Window
import UIAHandler
import wx
import tones

import controlTypes
import ui
import tones
from logHandler import log



# Constants
MODULE_UNKNOWN = "Unknown"
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


# For convenience.
ExtraUIAEvents = {
#    UIAHandler.UIA_AutomationFocusChangedEventId: "UIA_AutomationFocusChanged",
    UIAHandler.UIA_Selection_InvalidatedEventId: "UIA_selectionInvalidated"

}




class AppModule(appModuleHandler.AppModule):

    def __init__(self, *args, **kwargs):
        super(AppModule, self).__init__(*args, **kwargs)
        # Add a series of events instead of doing it one at a time.
        # Some events are only available in a specific build range and/or while a specific version of IUIAutomation interface is in use.
        #log.debug("SpcsAdm: adding additional events")
        for event, name in ExtraUIAEvents.items():
            if event not in UIAHandler.UIAEventIdsToNVDAEventNames:
                UIAHandler.UIAEventIdsToNVDAEventNames[event] = name
                UIAHandler.handler.clientObject.addAutomationEventHandler(event,UIAHandler.handler.rootElement,UIAHandler.TreeScope_Subtree,UIAHandler.handler.baseCacheRequest,UIAHandler.handler)
                #log.debug("SpcsAdm: added event ID %s, assigned to %s"%(event, name))
                
        confspec = {
            'debugMode': 'boolean(default=false)',
            'sayNumGridRows': 'boolean(default=false)'
        }
        config.conf.spec['VismaAdministration'] = confspec
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(VismaAdministrationSettingsPanel)


    def event_NVDAObject_init(self, obj):
        if not isinstance(obj, Window):
            pass
        if     obj.windowClassName == "SafGrid":
            if  obj.windowControlID == 7407:
                obj.name = 'Företagslista'
            elif     obj.windowControlID == 21346:
                obj.name = u"Lista"
            elif     obj.windowControlID == 21516:
                obj.name = "Artiklar"
        elif     obj.windowClassName == "SysTabControl32" and obj.windowControlID == 59648:
            obj.name = "Detaljsidor"
        # Artikelkort, flik 1
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22612:
            obj.name = "Artikelnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22613:
            obj.name = "Benämning"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22614:
            obj.name = "Annan benämning"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22615:
            obj.name = "Kortnamn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22616:
            obj.name = "Artikelgrupp"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22619:
            obj.name = "Enhet"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22617:
            obj.name = "Konteringskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22709:
            obj.name = "Streckkod"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 21657:
            obj.name = "Streckkodstyp"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22620:
            obj.name = "Sorteringsbegrepp"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22643        :
            obj.name = "Exportkod"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22641:
            obj.name = "Restnotera ej"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22618:
            obj.name = "Lagervara"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22652:
            obj.name = "Utgående"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 25050:
            obj.name = "Inaktiv"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 23606:
            obj.name = "Webshopartikel"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22708:
            obj.name = "Dokument"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22644:
            obj.name = "Anmärkning 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22645:
            obj.name = "Anmärkning 2"
        # Artikelkort, flik 2
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24279:
            obj.name = "Ursprungslandskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22642:
            obj.name = "Ursprungsland"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22628:
            obj.name = "Lagerplats"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22629:
            obj.name = "Vikt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22630:
            obj.name = "Volym"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22657:
            obj.name = "Lagervärdet är inaktivt"
        # Artikelkort, flik 3
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22627:
            obj.name = "Normal leveranstid"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22631:
            obj.name = "Huvudleverantör"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22653:
            obj.name = "Kalkylpris, inpris"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22654:
            obj.name = "Kalkylpris, frakt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22655:
            obj.name = "Kalkylpris, övrigt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22623:
            obj.name = "Beställningspunkt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22625:
            obj.name = "Maxnivå"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22624:
            obj.name = "Minsta beställning"
        # Artikelkort, flik 6
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 22673:
            obj.name = "Spårning"
        # Orderkort, flik 1
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22308:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Ordernummer"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Fakturanummer"
            
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22332:
            obj.name = "Kundnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22333:
            obj.name = "Namn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22344:
            obj.name = "Er referens"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22346:
            obj.name = "Vår referens"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23613:
            obj.name = "Kontraktsnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22362:
            obj.name = "Betalningsvillkor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22363:
            obj.name = "Leveransvillkor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22364:
            obj.name = "Leveranssätt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23658:
            obj.name = "Speditör"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22385:
            obj.name = "Resultatenhet"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22384:
            obj.name = "Projekt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22316:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('offert'):
                obj.name = "Offertdatum"
            elif wndtxt.startswith('order'):
                obj.name = "Orderdatum"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Fakturadatum"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22317:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('offert'):
                obj.name = "Giltig till och med"
            elif wndtxt.startswith('order'):
                obj.name = "Leveransdatum"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Förfallodatum"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22331:
            obj.name = "Ert referensnummer" # På offert
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23631:
            obj.name = "Leveransdatum" # På faktura
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22345:
            obj.name = "Ert ordernummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24346:
            obj.name = "Referenskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22359:
            obj.name = "Fakturarabatt i procent"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22377:
            obj.name = "Frakt"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22356:
            obj.name = "Inklusive moms"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22379:
            obj.name = "Expeditionsavgift"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22313:
            obj.name = "Ej klar"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22318:
            obj.name = "Levererad"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22315:
            obj.name = "Makulerad"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 21306:
            obj.name = "Journalförd"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22314:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Ordererkännande utskrivet"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Faktura utskriven"

        elif     obj.windowClassName == "Button" and obj.windowControlID == 22319:
            obj.name = "Plocklista utskriven"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22320:
            obj.name = "Följesedel utskriven"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 21196:
            obj.name = "Restorder"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 21229:
            obj.name = "Fakturerad"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24932:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Ordererkännande, utskriftsval"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Faktura, utskriftsval"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24986:
            obj.name = "Följesedel"
        # Orderkort, flik 2
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22334:
            obj.name = "Postadress"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22543:
            obj.name = "Postadress rad 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23733:
            obj.name = "GLN"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22335:
            obj.name = "Postnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22336:
            obj.name = "Ort"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23596:
            obj.name = "Landskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22337:
            obj.name = "Land"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22410:
            obj.name = "VAT-nummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22347:
            obj.name = "Distrikt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22348:
            obj.name = "Säljare"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22361:
            obj.name = "Rabattavtal"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22358:
            obj.name = "Prislista"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22349:
            obj.name = "Språk"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22350:
            obj.name = "Valuta"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22351:
            obj.name = "Kurs/enhet"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22338:
            obj.name = "Avvikande leveransadress, Namn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22339:
            obj.name = "Avvikande leveransadress, rad 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22545:
            obj.name = "Avvikande leveransadress, rad 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23736:
            obj.name = "Avvikande leveransadress, GLN"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22340:
            obj.name = "Avvikande leveransadress, Postnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22341:
            obj.name = "Avvikande leveransadress, Ort"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23597:
            obj.name = "Avvikande leveransadress, Landskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22342:
            obj.name = "Avvikande leveransadress, Land"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22355:
            obj.name = "Export"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22354:
            obj.name = "Räntefakturering"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22352:
            obj.name = "Betalningspåminnelse"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22353:
            obj.name = "Påminnelseavgift"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22357:
            obj.name = "Restnotera ej"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22411:
            obj.name = "Samlingsfakturera"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 24122:
            obj.name = "Överför adress till beställning"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 25083:
            obj.name = "Hämtad - Visma Webbfakturering"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22365:
            obj.name = "Text 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22366:
            obj.name = "Text 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22367:
            obj.name = "Text 3"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 25028:
            obj.name = "AutoInvoice-bilaga"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22584:
            obj.name = "Spara text"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22547:
            obj.name = "EU Periodisk sammanställning"
        # Orderkort, flik 3
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22378:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Momskod"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Momskod, Frakt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22380:
            obj.name = "Momskod, Expeditionsavgift"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22369:
            obj.name = "Procentsats, momskod 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22370:
            obj.name = "Procentsats, momskod 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22371:
            obj.name = "Procentsats, momskod 3"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22416:
            obj.name = "Procentsats, momskod 4"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 22548:
            obj.name = "Mellanman i trepartshandel EU"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22326:
            obj.name = "Ränta i procent"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22327:
            obj.name = "Konto för kundfordran"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22329:
            obj.name = "Konto för expeditionsavgift"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22328:
            obj.name = "Konto för frakt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22404:
            obj.name = "Konto för momskod 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22405:
            obj.name = "Konto för momskod 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22406:
            obj.name = "Konto för momskod 3"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22419:
            obj.name = "Konto för momskod 4"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22407:
            obj.name = "Konto för avrundning"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22408:
            obj.name = "Given rabatt"
        # Orderkort, flik 4 - Spårning
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22309:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Offertnummer / Datum"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Ordernummer / Datum"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22310:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Fakturanummer / Datum"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Avtalsnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22311:
            wndtxt = obj.parent.parent.parent.parent.parent.parent.windowText.lower()
            if wndtxt.startswith('order'):
                obj.name = "Restorder / Datum"
            elif wndtxt.startswith('kundfakt'):
                obj.name = "Journalnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22312:
            obj.name = "Ursprungsorder / Datum"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 21274:
            obj.name = "Offertnummer / Datum"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22414:
            obj.name = "Verifikationsnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24044:
            obj.name = "F U-nummer"
        # Fakturakort, flik 1
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22312:
            obj.name = "Ursprungsorder / Datum"
        # Leverantörskort, flik 1
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23179:
            obj.name = "Leverantörsnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23180:
            obj.name = "Namn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23191:
            obj.name = "Organisationsnummer"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 15162:
            obj.name = "Hämta uppgifter"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23181:
            obj.name = "Kortnamn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23182:
            obj.name = "Postadress"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23221:
            obj.name = "Postadress rad 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23734:
            obj.name = "GLN"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23183:
            obj.name = "Besöksadress"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23184:
            obj.name = "Postnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23185:
            obj.name = "Ort"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23210:
            obj.name = "Landskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23186:
            obj.name = "Land"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23187:
            obj.name = "Telefon"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23188:
            obj.name = "Telefon 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23189:
            obj.name = "Fax"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23787:
            obj.name = "Referenskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23190:
            obj.name = "Er referens"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 20748:
            obj.name = "Mobiltelefon"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 24921:
            obj.name = "Skicka SMS"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23228:
            obj.name = "E-post"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23229:
            obj.name = "WWW"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23209:
            obj.name = "Vår referens"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23206:
            obj.name = "Sorteringsbegrepp"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 25049:
            obj.name = "Inaktiv"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 26256:
            obj.name = "Autogiro"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 23230:
            obj.name = "Kräver OCR / Referenskod"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 23220:
            obj.name = "Utlandsbetalningar"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23196:
            obj.name = "Vårt kundnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23194:
            obj.name = "Bankgiro"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23195:
            obj.name = "PlusGiro"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 21770:
            obj.name = "Konteringsmall"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24824:
            obj.name = "Resultatenhet"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24825:
            obj.name = "Projekt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23197:
            obj.name = "Kreditgräns"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23193:
            obj.name = "Språk"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23192:
            obj.name = "Valuta"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23198:
            obj.name = "Betalningsvillkor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23207:
            obj.name = "Leveransvillkor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23208:
            obj.name = "Leveranssätt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23232:
            obj.name = "Dokument"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23204:
            obj.name = "Anmärkning 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23205:
            obj.name = "Anmärkning 2"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 24920:
            obj.name = "Visma Kreditupplysning"
        # Leverantörskort, flik 2
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23213:
            obj.name = "BIC"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23214:
            obj.name = "Bankkod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23215:
            obj.name = "Bank"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23216:
            obj.name = "Adress"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23218:
            obj.name = "Postnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23219:
            obj.name = "Ort"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 26515:
            obj.name = "Landskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23211:
            obj.name = "Betalningskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23212:
            obj.name = "Mottagarnummer"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 23590:
            obj.name = "Avgiftskod"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 23231:
            obj.name = "Betalas med valutakonto / ficka"
        # Leverantörskort, flik 4
        elif     obj.windowClassName == "Button" and obj.windowControlID == 24974:
            obj.name = "Kontrollera om leverantören kan ta emot e-beställning"
        elif     obj.windowClassName == "Button" and obj.windowControlID == 24990:
            obj.name = "Kopia till leverantör"
        # Kund, flik 1
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22423:
            obj.name = "Kundnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22424:
            obj.name = "Namn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22444:
            obj.name = "Organisationsnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22426:
            obj.name = "Postadress rad 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22542:
            obj.name = "Postadress rad 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23732:
            obj.name = "GLN"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22427:
            obj.name = "Besöksadress"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22428:
            obj.name = "Postnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22429:
            obj.name = "Ort"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23600:
            obj.name = "Landskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22430:
            obj.name = "Land"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22483:
            obj.name = "VAT-nummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22431:
            obj.name = "Telefon"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22432:
            obj.name = "Telefon 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23628:
            obj.name = "Telefon 3"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22433:
            obj.name = "Fax"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22425:
            obj.name = "Kortnamn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23808:
            obj.name = "Referenskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22443:
            obj.name = "Referens"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 20747:
            obj.name = "Mobiltelefon"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22485:
            obj.name = "E-post"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22434:
            obj.name = "Avvikande Leveransadress, namn"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22435:
            obj.name = "Avvikande Leveransadress, rad 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22544:
            obj.name = "Avvikande Leveransadress, rad 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23735:
            obj.name = "Avvikande Leveransadress, GLN"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22436:
            obj.name = "Avvikande Leveransadress, besöksadress"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22437:
            obj.name = "Avvikande Leveransadress, postnummer"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22438:
            obj.name = "Avvikande Leveransadress, ort"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23601:
            obj.name = "Avvikande Leveransadress, landskod"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22439:
            obj.name = "Avvikande Leveransadress, land"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22440:
            obj.name = "Avvikande Leveransadress, telefon"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22441:
            obj.name = "Avvikande Leveransadress, telefon 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23626:
            obj.name = "Avvikande Leveransadress, telefon 3"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22442:
            obj.name = "Avvikande Leveransadress, fax"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22446:
            obj.name = "Kundkategori"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22447:
            obj.name = "Distrikt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22448:
            obj.name = "Säljare"
        # Kund, flik 2
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22450:
            obj.name = "Betalningsvillkor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22449:
            obj.name = "Leveransvillkor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22451:
            obj.name = "Leveranssätt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 23737:
            obj.name = "Speditör"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22457:
            obj.name = "Språk"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22456:
            obj.name = "Valuta"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24821:
            obj.name = "Resultatenhet"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 24822:
            obj.name = "Projekt"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22481:
            obj.name = "Sorteringsbegrepp"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22479:
            obj.name = "Bankgiro"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22480:
            obj.name = "PlusGiro"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22467:
            obj.name = "Kreditgräns i kronor"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22458:
            obj.name = "Rabattavtal"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22486:
            obj.name = "Dokument"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22465:
            obj.name = "Anmärkning 1"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 22466:
            obj.name = "Anmärkning 2"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 26461:
            obj.name = "Fastighets-/lägenhetsbeteckning"
        elif     obj.windowClassName == "Edit" and obj.windowControlID == 26462:
            obj.name = "Bostadsrättsföreningens organisationsnummer"
        # Kund, flik 3
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24838:
            obj.name = "Information på offert"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24842:
            obj.name = "Information på order"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24846:
            obj.name = "Information på följesedel"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24849:
            obj.name = "Information på extra orderdokument"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24852:
            obj.name = "Information på faktura"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24856:
            obj.name = "Information på extra fakturadokuments"
        elif     obj.windowClassName == "ComboBox" and obj.windowControlID == 24859:
            obj.name = "Information på påminnelse"




            

        else:
            if config.conf['VismaAdministration']['debugMode']:
                if obj.windowClassName == "Button" or obj.windowClassName == "Edit" or obj.windowClassName == "ComboBox":
                    try:
                        obj.name = "%s, %s, %d, %s" % ( obj.name, obj.displayText, obj.windowControlID, obj.windowClassName )
                    except:
                        obj.name = "%d, %s" % ( obj.windowControlID, obj.windowClassName )

            
    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        if obj.windowClassName == "SafGrid" or obj.windowClassName == "AfxWnd140s":
            clsList.insert(0, VismaSafGrid)

#    def ReadControlInfo(self):
#        ui.message("%s" % self.windowControlID)

    def getCurrentVismaModule(self):
        try:
            wnd = api.getFocusObject()
            if wnd is None:
                return MODULE_UNKNOWN
            while wnd is not None and isinstance(wnd, Window) and not wnd.windowClassName.startswith('Afx:'):
                wnd = wnd.parent
                if wnd is None:
                    return MODULE_UNKNOWN
                #log.debug(wnd.windowClassName)
                #ui.message(wnd.windowClassName)
            wndtxt = wnd.windowText.lower()
            if wndtxt.startswith('order'):
                return MODULE_ORDER
            elif wndtxt.startswith('kundfakt'):
                return MODULE_INVOICE
            elif wndtxt.startswith('offerter'):
                return MODULE_OFFER
            elif wndtxt.startswith('kunder'):
                return MODULE_CUSTOMER
            elif wndtxt.startswith('leverantörer'):
                return MODULE_SUPPLIER
            elif wndtxt.startswith('artiklar'):
                return MODULE_ARTICLE
            elif wndtxt.startswith('beställningar'):
                return MODULE_BOOKINGS
            elif wndtxt.startswith('inkommande följ'):
                return MODULE_DELIVERYNOTE
            elif wndtxt.startswith('leverantörsfakturor'):
                return MODULE_SUPPLIERINVO
            elif wndtxt.startswith('kontakter'):
                return MODULE_CONTACTS
            elif wndtxt.startswith('avtalsmall'):
                return MODULE_AGREEMENT_TEMPLATE
            elif wndtxt.startswith('avtal'):
                return MODULE_AGREEMENT
            elif wndtxt.startswith('försäljningsprislis'):
                return MODULE_PRICELIST
            elif wndtxt.startswith('manuella inleveranser'):
                return MODULE_MANUALDELIVERYIN
            elif wndtxt.startswith('manuella utleveranser'):
                return MODULE_MANUALDELIVERYOU
            elif wndtxt.startswith('inventering'):
                return MODULE_INVENTORY
            elif wndtxt.startswith('prisinläsning'):
                return MODULE_PRICEIMPORT
            elif wndtxt.startswith('prisomräkning kalkyl'):
                return MODULE_PRICERECALCULATION_ESTIMATEDPRICES
            elif wndtxt.startswith('prisomräkning lev'):
                return MODULE_PRICERECALCULATION_SUPPLIE
            elif wndtxt.startswith('ingående balans'):
                return MODULE_INGOINGBALLANCE
            elif wndtxt.startswith('verifikationer'):
                return MODULE_VERIFICATIONS
            elif wndtxt.startswith('resultatbudg'):
                return MODULE_RESULTBUDGET
            elif wndtxt.startswith('resultatprognos'):
                return MODULE_RESULTPROGNOZE
            elif wndtxt.startswith('balansbudget'):
                return MODULE_BALLANCEBUDGET
            elif wndtxt.startswith('balansprognos'):
                return MODULE_BALLANCEPROGNOZE

            return "Fönstertext " + wndtxt
        except Exception as e:
            ui.message("Fel: %s" % e)


    def doReadVismaCommands(self):
        try:
            modul = self.getCurrentVismaModule()
            ui.message(modul)
            fn = os.path.dirname(os.path.abspath(__file__)) + "\\shortcuts.txt"
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
        except Exception as e:
            ui.message("Fel: %s" % e)


    def script_readVismaCommands(self, gesture):
        self.doReadVismaCommands()

    __gestures = {
        "kb:NVDA+h": "readVismaCommands"
    }


class VismaSafGrid(UIA):


    def event_gainFocus(self):
        try:
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
                        elif coltxt == "Kol 1":
                            coltxt = "Mapp"
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
                        valtxt = valpat.CurrentValue;
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
                        valtxt = valpat.CurrentValue;
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
        #"kb:NVDA+i": "readControlInfo"
        #"kb:downarrow": "readGridSelection",
        #"kb:uparrow": "readGridSelection",
        #"kb:leftarrow": "readGridSelection",
        #"kb:rightarrow": "readGridSelection"
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

