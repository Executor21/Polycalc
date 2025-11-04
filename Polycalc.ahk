/*
Script: Polycalc
Συγγραφέας: Tasos
Έτος: 2025
MIT License
Copyright (c) 2025 Tasos
*/
#Requires AutoHotkey v2.0+
#SingleInstance Force
SetWorkingDir(A_ScriptDir)
FileEncoding("UTF-16")

MAIN_PROGRAM()

MAIN_PROGRAM() {
global MyGui, configFile, currentEditingApt, ListView, TotalExpenses
global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl
global CommonTotalCtrl, ElevatorTotalCtrl, HeatingTotalCtrl, PrintingTotalCtrl, ReserveTotalCtrl, GrandTotalCtrl
global AptsListView, AptNameCtrl, AptOwnerCtrl, AptCommonCtrl, AptElevatorCtrl, AptHeatingCtrl
global AptSelectDropdown
global AptCommonBaseCtrl, AptCommonPercentCtrl, AptCommonResultCtrl
global AptElevatorBaseCtrl, AptElevatorPercentCtrl, AptElevatorResultCtrl
global AptHeatingBaseCtrl, AptHeatingPercentCtrl, AptTotalHeatingPercentCtrl, AptHeatingResultCtrl
global AptPrintingBaseCtrl, AptPrintingPercentCtrl, AptPrintingResultCtrl
global AptReserveBaseCtrl, AptReservePercentCtrl, AptReserveResultCtrl, AptTotalResultCtrl
global TabCtrl, StatusBar

TraySetIcon("Shell32.dll", 44)
configFile := "Polykatoikia_Data.ini"
currentEditingApt := ""
TotalExpenses := Map()

; Δημιουργία παραθύρου με σκούρο θέμα
MyGui := Gui(, "Polycalc")
MyGui.SetFont("s10", "Segoe UI")
MyGui.BackColor := "0xF0F0F0"
MyGui.Opt("-Resize +MaximizeBox +MinimizeBox")

; Status Bar
StatusBar := MyGui.AddStatusBar(, "Έτοιμο | Έκδοση: v1.0")

; Δημιουργία καρτέλας
TabCtrl := MyGui.AddTab3("x10 y10 w1050 h660", ["📊 Δαπάνες", "🏢 Διαμερίσματα", "📈 Αναλυτικά"])

; ═══════════════════════════════════════════════════════════
; ΚΑΡΤΕΛΑ 1: ΔΑΠΑΝΕΣ
; ═══════════════════════════════════════════════════════════
TabCtrl.UseTab(1)

MyGui.SetFont("s12 Bold", "Segoe UI")
MyGui.Add("Text", "x20 y40 w1000 h30 Center c0x2C5F2D BackgroundWhite", "ΕΙΣΑΓΩΓΗ ΔΑΠΑΝΩΝ ΠΟΛΥΚΑΤΟΙΚΙΑΣ")
MyGui.SetFont("s10 Norm", "Segoe UI")

; ═══ ΑΡΙΣΤΕΡΗ ΣΤΗΛΗ - ΕΙΣΑΓΩΓΗ ΔΑΠΑΝΩΝ ═══
; GroupBox για Κοινόχρηστα
CommonGroup := MyGui.Add("GroupBox", "x20 y80 w420 h210", "ΚΟΙΝΟΧΡΗΣΤΑ")
CommonGroup.SetFont("s10 Bold")

MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y105 w180 h23", "💧 Καθαριότητα (€):")
CleanCtrl := MyGui.Add("Edit", "x230 y105 w190 h25 Background0xFFFFFF", "0,00")

MyGui.Add("Text", "x40 y135 w180 h23", "⚡ Ηλ. Ρεύμα (€):")
ElectricityCtrl := MyGui.Add("Edit", "x230 y135 w190 h25 Background0xFFFFFF", "0,00")

MyGui.Add("Text", "x40 y165 w180 h23", "🚰 Νερό (€):")
WaterCtrl := MyGui.Add("Edit", "x230 y165 w190 h25 Background0xFFFFFF", "0,00")

MyGui.Add("Text", "x40 y195 w180 h23", "🔥 Πυρασφάλεια (€):")
FireCtrl := MyGui.Add("Edit", "x230 y195 w190 h25 Background0xFFFFFF", "0,00")

MyGui.Add("Text", "x40 y225 w180 h23", "🌳 Κηπουρός (€):")
GardenerCtrl := MyGui.Add("Edit", "x230 y225 w190 h25 Background0xFFFFFF", "0,00")

MyGui.Add("Text", "x40 y255 w180 h23", "💼 Άλλα έξοδα (€):")
OtherCtrl := MyGui.Add("Edit", "x230 y255 w190 h25 Background0xFFFFFF", "0,00")

; GroupBox για Ασανσέρ
ElevatorGroup := MyGui.Add("GroupBox", "x20 y300 w420 h55", "ΑΣΑΝΣΕΡ") 
ElevatorGroup.SetFont("s10 Bold")
MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y320 w180 h23", "🛗 Ασανσέρ (€):")
ElevatorCtrl := MyGui.Add("Edit", "x230 y320 w190 h25 Background0xFFFFFF", "0,00")

; GroupBox για Θέρμανση
HeatingGroup := MyGui.Add("GroupBox", "x20 y365 w420 h55", "ΘΕΡΜΑΝΣΗ")
HeatingGroup.SetFont("s10 Bold")
MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y385 w180 h23", "🔥 Θέρμανση (€):")
HeatingCtrl := MyGui.Add("Edit", "x230 y385 w190 h25 Background0xFFFFFF", "0,00")

; GroupBox για Έκδοση
PrintingGroup := MyGui.Add("GroupBox", "x20 y430 w420 h55", "ΕΚΔΟΣΗ ΚΟΙΝΟΧΡΗΣΤΩΝ") 
PrintingGroup.SetFont("s10 Bold")
MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y450 w180 h23", "🖨️ Έκδοση (€):")
PrintingCtrl := MyGui.Add("Edit", "x230 y450 w190 h25 Background0xFFFFFF", "0,00")

; GroupBox για Αποθεματικό
ReserveGroup := MyGui.Add("GroupBox", "x20 y495 w420 h55", "ΑΠΟΘΕΜΑΤΙΚΟ")
ReserveGroup.SetFont("s10 Bold")
MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y515 w180 h23", "💰 Αποθεματικό (€):")
ReserveCtrl := MyGui.Add("Edit", "x230 y515 w190 h25 Background0xFFFFFF", "0,00")

; Κουμπιά Ενεργειών
MyGui.SetFont("s10 Bold")
ClearBtn := MyGui.Add("Button", "x20 y565 w130 h35", "🗑️ ΚΑΘΑΡΙΣΜΟΣ")  
SaveDataBtn := MyGui.Add("Button", "x160 y565 w130 h35", "💾 ΑΠΟΘΗΚΕΥΣΗ")
LoadDataBtn := MyGui.Add("Button", "x310 y565 w130 h35", "📂 ΦΟΡΤΩΣΗ")
InfoBtn := MyGui.Add("Button", "x20 y610 w420 h35", "ℹ️ ΠΛΗΡΟΦΟΡΙΕΣ")  

; ═══ ΔΕΞΙΑ ΣΤΗΛΗ - ΑΠΟΤΕΛΕΣΜΑΤΑ ═══
ResultsGroup := MyGui.Add("GroupBox", "x460 y80 w580 h520", "ΣΥΝΟΛΙΚΑ ΑΠΟΤΕΛΕΣΜΑΤΑ")
ResultsGroup.SetFont("s11 Bold")

MyGui.SetFont("s10 Norm")
MyGui.Add("Text", "x480 y120 w250 h25", "ΣΥΝΟΛΟ ΚΟΙΝΟΧΡΗΣΤΩΝ:")
CommonTotalCtrl := MyGui.Add("Edit", "x750 y120 w270 h30 ReadOnly Background0xE8F5E9 Center", "0.00 €")
CommonTotalCtrl.SetFont("s11 Bold c0x1B5E20")

MyGui.Add("Text", "x480 y160 w250 h25", "ΣΥΝΟΛΟ ΑΣΑΝΣΕΡ:")
ElevatorTotalCtrl := MyGui.Add("Edit", "x750 y160 w270 h30 ReadOnly Background0xE3F2FD Center", "0.00 €")
ElevatorTotalCtrl.SetFont("s11 Bold c0x0D47A1")

MyGui.Add("Text", "x480 y200 w250 h25", "ΣΥΝΟΛΟ ΘΕΡΜΑΝΣΗΣ:")
HeatingTotalCtrl := MyGui.Add("Edit", "x750 y200 w270 h30 ReadOnly Background0xFFF3E0 Center", "0.00 €")
HeatingTotalCtrl.SetFont("s11 Bold c0xE65100")

MyGui.Add("Text", "x480 y240 w250 h25", "ΣΥΝΟΛΟ ΕΚΔΟΣΗΣ:")
PrintingTotalCtrl := MyGui.Add("Edit", "x750 y240 w270 h30 ReadOnly Background0xF3E5F5 Center", "0.00 €")
PrintingTotalCtrl.SetFont("s11 Bold c0x6A1B9A")

MyGui.Add("Text", "x480 y280 w250 h25", "ΣΥΝΟΛΟ ΑΠΟΘΕΜΑΤΙΚΟΥ:")
ReserveTotalCtrl := MyGui.Add("Edit", "x750 y280 w270 h30 ReadOnly Background0xFFF9C4 Center", "0.00 €")
ReserveTotalCtrl.SetFont("s11 Bold c0xF57F17")

MyGui.Add("Text", "x480 y340 w540 h2 0x10")
MyGui.SetFont("s12 Bold")
MyGui.Add("Text", "x480 y355 w250 h30", "ΓΕΝΙΚΟ ΣΥΝΟΛΟ:")
GrandTotalCtrl := MyGui.Add("Edit", "x750 y355 w270 h40 ReadOnly Background0xFFEBEE Center", "0.00 €")
GrandTotalCtrl.SetFont("s14 Bold c0xC62828")

; Πληροφοριακό κείμενο
MyGui.SetFont("s9 Italic", "Segoe UI")
MyGui.Add("Text", "x480 y400 w540 h150", "💡 ΣΥΜΒΟΥΛΗ:`n`nΤα ποσά υπολογίζονται αυτόματα καθώς πληκτρολογείτε.`n`nΜην ξεχάσετε να αποθηκεύσετε τις αλλαγές σας!`n`nΓια περισσότερες πληροφορίες, πατήστε το κουμπί 'ΠΛΗΡΟΦΟΡΙΕΣ'.")

; ═══════════════════════════════════════════════════════════
; ΚΑΡΤΕΛΑ 2: ΔΙΑΜΕΡΙΣΜΑΤΑ
; ═══════════════════════════════════════════════════════════
TabCtrl.UseTab(2)

MyGui.SetFont("s12 Bold", "Segoe UI")
MyGui.Add("Text", "x20 y40 w1000 h30 Center c0x1565C0 BackgroundWhite", "ΔΙΑΧΕΙΡΙΣΗ ΔΙΑΜΕΡΙΣΜΑΤΩΝ")
MyGui.SetFont("s10 Norm", "Segoe UI")

; Πίνακας διαμερισμάτων με βελτιωμένη εμφάνιση
MyGui.Add("Text", "x20 y90 w1000 h25 Center Background0x1565C0 cWhite", "📋 ΛΙΣΤΑ ΔΙΑΜΕΡΙΣΜΑΤΩΝ")
AptsListView := MyGui.AddListView("x20 y110 w1000 h300 Background0xFFFFFF Grid", [
    "Διαμέρισμα", 
    "Ιδιοκτήτης", 
    "Κοινοχ. ΧΙΛ.", 
    "Ασανσέρ ΧΙΛ.", 
    "Θέρμανση ΧΙΛ.", 
    "Εκτύπωση %",
    "Πληρωτέο Ποσό"
])

; Κουμπιά διαχείρισης με icons
MyGui.SetFont("s9 Bold")
AddAptBtn := MyGui.Add("Button", "x20 y420 w155 h35", "➕ ΠΡΟΣΘΗΚΗ") 
EditAptBtn := MyGui.Add("Button", "x185 y420 w155 h35", "✏️ ΕΠΕΞΕΡΓΑΣΙΑ")
DeleteAptBtn := MyGui.Add("Button", "x350 y420 w155 h35", "🗑️ ΔΙΑΓΡΑΦΗ")
CalculateAptsBtn := MyGui.Add("Button", "x515 y420 w155 h35", "🧮 ΥΠΟΛΟΓΙΣΜΟΣ")
SaveAptsBtn := MyGui.Add("Button", "x680 y420 w155 h35", "💾 ΑΠΟΘΗΚΕΥΣΗ")
LoadAptsBtn := MyGui.Add("Button", "x865 y420 w155 h35", "📂 ΦΟΡΤΩΣΗ")

; GroupBox για εισαγωγή νέου διαμερίσματος
MyGui.SetFont("s10 Norm")
NewAptGroup := MyGui.Add("GroupBox", "x20 y465 w1000 h160", "ΕΙΣΑΓΩΓΗ/ΕΠΕΞΕΡΓΑΣΙΑ ΔΙΑΜΕΡΙΣΜΑΤΟΣ")
NewAptGroup.SetFont("s10 Bold")

MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y490 w120 h23", "🏠 Διαμέρισμα:")
AptNameCtrl := MyGui.Add("Edit", "x170 y490 w150 h25 Background0xFFFFFF", "")

MyGui.Add("Text", "x340 y490 w120 h23", "👤 Ιδιοκτήτης:")
AptOwnerCtrl := MyGui.Add("Edit", "x470 y490 w250 h25 Background0xFFFFFF", "")

MyGui.Add("Text", "x40 y525 w120 h23", "📊 Κοινοχ. ΧΙΛ.:")
AptCommonCtrl := MyGui.Add("Edit", "x170 y525 w100 h25 Background0xFFFFFF", "0")

MyGui.Add("Text", "x290 y525  w120 h23", "🛗 Ασανσέρ ΧΙΛ.:")
AptElevatorCtrl := MyGui.Add("Edit", "x420 y525 w100 h25 Background0xFFFFFF", "0")

MyGui.Add("Text", "x540 y525 w120 h23", "🔥 Θέρμανση ΧΙΛ.:")
AptHeatingCtrl := MyGui.Add("Edit", "x670 y525 w100 h25 Background0xFFFFFF", "0")

MyGui.SetFont("s9 Bold")
AddAptFinalBtn := MyGui.Add("Button", "x40 y565 w200 h40", "✅ ΠΡΟΣΘΗΚΗ ΔΙΑΜΕΡΙΣΜΑΤΟΣ") 
SaveTableBtn := MyGui.Add("Button", "x250 y565 w200 h40", "📄 ΑΠΟΘΗΚΕΥΣΗ ΠΙΝΑΚΑ")
IniForPolyfundBtn := MyGui.Add("Button", "x460 y565 w200 h40", "💾 INI FOR POLYFUND")

; ═══════════════════════════════════════════════════════════
; ΚΑΡΤΕΛΑ 3: ΑΝΑΛΥΤΙΚΑ
; ═══════════════════════════════════════════════════════════
TabCtrl.UseTab(3)

MyGui.SetFont("s12 Bold", "Segoe UI")
MyGui.Add("Text", "x20 y40 w1000 h30 Center c0x2E7D32 BackgroundWhite", "ΑΝΑΛΥΤΙΚΑ ΕΞΟΔΑ ΑΝΑ ΔΙΑΜΕΡΙΣΜΑ")
MyGui.SetFont("s10 Norm", "Segoe UI")

; Επιλογή διαμερίσματος
SelectGroup := MyGui.Add("GroupBox", "x20 y80 w1000 h60", "ΕΠΙΛΟΓΗ ΔΙΑΜΕΡΙΣΜΑΤΟΣ")
SelectGroup.SetFont("s10 Bold")
MyGui.SetFont("s9 Norm")
MyGui.Add("Text", "x40 y115 w180 h25", "🏠 Επιλογή διαμερίσματος:")
AptSelectDropdown := MyGui.Add("DropDownList", "x230 y105 w400 h200")
MyGui.SetFont("s9 Bold")
SaveAptResultsBtn := MyGui.Add("Button", "x650 y102 w350 h30", "💾 ΑΠΟΘΗΚΕΥΣΗ ΣΕ ΑΡΧΕΙΟ")

; Αναλυτικά αποτελέσματα
DetailsGroup := MyGui.Add("GroupBox", "x20 y150 w1000 h410", "ΑΝΑΛΥΤΙΚΟΣ ΥΠΟΛΟΓΙΣΜΟΣ")
DetailsGroup.SetFont("s11 Bold")

MyGui.SetFont("s9 Norm")
; Κοινόχρηστα
MyGui.Add("Text", "x40 y177 w280 h23 Background0xE8F5E9", "📊 ΣΥΝΟΛΟ ΚΟΙΝΟΧΡΗΣΤΩΝ:")
AptCommonBaseCtrl := MyGui.Add("Edit", "x320 y175 w100 h25 ReadOnly Background0xF5F5F5 Center", "0.00 €")
MyGui.Add("Text", "x430 y177 w30 h23 Center", "×")
MyGui.Add("Text", "x460 y177 w50 h23", "ΧΙΛ:")
AptCommonPercentCtrl := MyGui.Add("Edit", "x510 y175 w60 h25 ReadOnly Background0xF5F5F5 Center", "0")
MyGui.Add("Text", "x580 y177 w50 h23 Center", "/1000")
MyGui.Add("Text", "x640 y177 w30 h23 Center", "=")
AptCommonResultCtrl := MyGui.Add("Edit", "x680 y175 w140 h25 ReadOnly Background0xE8F5E9 Center", "0.00 €")
AptCommonResultCtrl.SetFont("s10 Bold c0x1B5E20")

; Ασανσέρ
MyGui.Add("Text", "x40 y212 w280 h23 Background0xE3F2FD", "🛗 ΣΥΝΟΛΟ ΑΣΑΝΣΕΡ:")
AptElevatorBaseCtrl := MyGui.Add("Edit", "x320 y210 w100 h25 ReadOnly Background0xF5F5F5 Center", "0.00 €")
MyGui.Add("Text", "x430 y212 w30 h23 Center", "×")
MyGui.Add("Text", "x460 y212 w50 h23", "ΧΙΛ:")
AptElevatorPercentCtrl := MyGui.Add("Edit", "x510 y210 w60 h25 ReadOnly Background0xF5F5F5 Center", "0")
MyGui.Add("Text", "x580 y212 w50 h23 Center", "/1000")
MyGui.Add("Text", "x640 y212 w30 h23 Center", "=")
AptElevatorResultCtrl := MyGui.Add("Edit", "x680 y210 w140 h25 ReadOnly Background0xE3F2FD Center", "0.00 €")
AptElevatorResultCtrl.SetFont("s10 Bold c0x0D47A1")

; Θέρμανση
MyGui.Add("Text", "x40 y247 w280 h23 Background0xFFF3E0", "🔥 ΣΥΝΟΛΟ ΘΕΡΜΑΝΣΗΣ:")
AptHeatingBaseCtrl := MyGui.Add("Edit", "x320 y245 w100 h25 ReadOnly Background0xF5F5F5 Center", "0.00 €")
MyGui.Add("Text", "x430 y247 w30 h23 Center", "×")
MyGui.Add("Text", "x460 y247 w50 h23", "ΧΙΛ:")
AptHeatingPercentCtrl := MyGui.Add("Edit", "x510 y245 w60 h25 ReadOnly Background0xF5F5F5 Center", "0")
MyGui.Add("Text", "x580 y247 w20 h23 Center", "/")
AptTotalHeatingPercentCtrl := MyGui.Add("Edit", "x600 y245 w50 h25 ReadOnly Background0xF5F5F5 Center", "0")
MyGui.Add("Text", "x640 y247 w30 h23 Center", "=")
AptHeatingResultCtrl := MyGui.Add("Edit", "x680 y245 w140 h25 ReadOnly Background0xFFF3E0 Center", "0.00 €")
AptHeatingResultCtrl.SetFont("s10 Bold c0xE65100")

; Έκδοση
MyGui.Add("Text", "x40 y282 w280 h23 Background0xF3E5F5", "🖨️ ΣΥΝΟΛΟ ΕΚΔΟΣΗΣ:")
AptPrintingBaseCtrl := MyGui.Add("Edit", "x320 y280 w100 h25 ReadOnly Background0xF5F5F5 Center", "0.00 €")
MyGui.Add("Text", "x430 y282 w30 h23 Center", "×")
MyGui.Add("Text", "x460 y282 w50 h23", "%:")
AptPrintingPercentCtrl := MyGui.Add("Edit", "x510 y280 w60 h25 ReadOnly Background0xF5F5F5 Center", "0")
MyGui.Add("Text", "x580 y282 w50 h23 Center", "/100")
MyGui.Add("Text", "x640 y282 w30 h23 Center", "=")
AptPrintingResultCtrl := MyGui.Add("Edit", "x680 y280 w140 h25 ReadOnly Background0xF3E5F5 Center", "0.00 €")
AptPrintingResultCtrl.SetFont("s10 Bold c0x6A1B9A")

; Αποθεματικό
MyGui.Add("Text", "x40 y317 w280 h23 Background0xFFF9C4", "💰 ΣΥΝΟΛΟ ΑΠΟΘΕΜΑΤΙΚΟΥ:")
AptReserveBaseCtrl := MyGui.Add("Edit", "x320 y315 w100 h25 ReadOnly Background0xF5F5F5 Center", "0.00 €")
MyGui.Add("Text", "x430 y317 w30 h23 Center", "×")
MyGui.Add("Text", "x460 y317 w50 h23", "ΧΙΛ:")
AptReservePercentCtrl := MyGui.Add("Edit", "x510 y315 w60 h25 ReadOnly Background0xF5F5F5 Center", "0")
MyGui.Add("Text", "x580 y317 w50 h23 Center", "/1000")
MyGui.Add("Text", "x640 y317 w30 h23 Center", "=")
AptReserveResultCtrl := MyGui.Add("Edit", "x680 y315 w140 h25 ReadOnly Background0xFFF9C4 Center", "0.00 €")
AptReserveResultCtrl.SetFont("s10 Bold c0xF57F17")

; Γραμμή διαχωρισμού
MyGui.Add("Text", "x40 y360 w960 h2 0x10")

; Γενικό σύνολο
MyGui.SetFont("s11 Bold")
MyGui.Add("Text", "x40 y375 w280 h30 Background0xFFEBEE", "💵 ΓΕΝΙΚΟ ΣΥΝΟΛΟ:")
AptTotalResultCtrl := MyGui.Add("Edit", "x320 y375 w500 h30 ReadOnly Background0xFFCDD2 Center", "0.00 €")
AptTotalResultCtrl.SetFont("s14 Bold c0xC62828")

; Πληροφοριακό κείμενο
MyGui.SetFont("s9 Italic", "Segoe UI")
MyGui.Add("Text", "x40 y420 w960 h120", "📌 ΣΗΜΕΙΩΣΗ:`n`nΕπιλέξτε ένα διαμέρισμα από το dropdown μενού για να δείτε τα αναλυτικά έξοδά του.`n`nΟ υπολογισμός γίνεται αυτόματα με βάση τα χιλιοστά και τα ποσοστά που έχετε ορίσει.`n`nΜπορείτε να αποθηκεύσετε τα αποτελέσματα σε αρχείο κειμένου πατώντας το κουμπί 'ΑΠΟΘΗΚΕΥΣΗ ΣΕ ΑΡΧΕΙΟ'.")

; Επιστροφή στην πρώτη καρτέλα
TabCtrl.UseTab(1)

currentEditingApt := ""

; Σύνδεση συναρτήσεων για φιλτράρισμα εισαγωγής
CleanCtrl.OnEvent("Change", FilterNumericInput)
ElectricityCtrl.OnEvent("Change", FilterNumericInput)
WaterCtrl.OnEvent("Change", FilterNumericInput)
FireCtrl.OnEvent("Change", FilterNumericInput)
GardenerCtrl.OnEvent("Change", FilterNumericInput)
OtherCtrl.OnEvent("Change", FilterNumericInput)
ElevatorCtrl.OnEvent("Change", FilterNumericInput)
HeatingCtrl.OnEvent("Change", FilterNumericInput)
PrintingCtrl.OnEvent("Change", FilterNumericInput)
ReserveCtrl.OnEvent("Change", FilterNumericInput)

; Σύνδεση συναρτήσεων με τα κουμπιά
ClearBtn.OnEvent("Click", ClearData)
SaveDataBtn.OnEvent("Click", SaveData)
LoadDataBtn.OnEvent("Click", LoadData)
InfoBtn.OnEvent("Click", ShowInfo)
AddAptBtn.OnEvent("Click", AddApartment)
EditAptBtn.OnEvent("Click", EditApartment)
DeleteAptBtn.OnEvent("Click", DeleteApartment)
CalculateAptsBtn.OnEvent("Click", CalculateAndGoToTab3)
SaveAptsBtn.OnEvent("Click", SaveApartmentChanges)
SaveTableBtn.OnEvent("Click", SaveTableToFile)
LoadAptsBtn.OnEvent("Click", LoadApartments)
AddAptFinalBtn.OnEvent("Click", AddApartment)
SaveAptResultsBtn.OnEvent("Click", SaveResultsToFile)
IniForPolyfundBtn.OnEvent("Click", SaveIniForPolyfund)

; Σύνδεση για αυτόματο υπολογισμό
CleanCtrl.OnEvent("Change", AutoCalculateExpenses)
ElectricityCtrl.OnEvent("Change", AutoCalculateExpenses)
WaterCtrl.OnEvent("Change", AutoCalculateExpenses)
FireCtrl.OnEvent("Change", AutoCalculateExpenses)
GardenerCtrl.OnEvent("Change", AutoCalculateExpenses)
OtherCtrl.OnEvent("Change", AutoCalculateExpenses)
ElevatorCtrl.OnEvent("Change", AutoCalculateExpenses)
HeatingCtrl.OnEvent("Change", AutoCalculateExpenses)
PrintingCtrl.OnEvent("Change", AutoCalculateExpenses)
ReserveCtrl.OnEvent("Change", AutoCalculateExpenses)
AptSelectDropdown.OnEvent("Change", AutoCalculateApartment)

; Εκκίνηση
LoadData()
MyGui.Show("w1080 h720")
LoadApartments()
UpdateApartmentDropdown()
CalculateTotalExpenses()
UpdatePrintingPercentages()

StatusBar.SetText("✅ Το πρόγραμμα είναι έτοιμο προς χρήση!")
}

; ═══════════════════════════════════════════════════════════
; ΒΟΗΘΗΤΙΚΕΣ ΣΥΝΑΡΤΗΣΕΙΣ
; ═══════════════════════════════════════════════════════

GetApartmentCount() {
    global configFile
    count := 0
    if !FileExist(configFile)
        return 0
    sections := IniRead(configFile)
    if (sections = "ERROR")
        return 0
    loop parse, sections, "`n" {
        section := Trim(A_LoopField)
        if (InStr(section, "Apartment_"))
            count++
    }
    return count
}

UpdatePrintingPercentages() {
    global configFile, StatusBar
    apartmentCount := GetApartmentCount()
    if (apartmentCount = 0)
        return
    printingPercent := Format("{:.2f}", 100 / apartmentCount)
    sections := IniRead(configFile)
    if (sections = "ERROR")
        return
    loop parse, sections, "`n" {
        section := Trim(A_LoopField)
        if (InStr(section, "Apartment_"))
            IniWrite(printingPercent, configFile, section, "PrintingPercent")
    }
    StatusBar.SetText("✅ Ποσοστά εκτύπωσης ενημερώθηκαν!")
}

UpdateApartmentDropdown() {
    global configFile, AptSelectDropdown
    AptSelectDropdown.Delete()
    if !FileExist(configFile)
        return
    sections := IniRead(configFile)
    if (sections = "ERROR")
        return
    apartmentNames := []
    loop parse, sections, "`n" {
        section := Trim(A_LoopField)
        if (InStr(section, "Apartment_")) {
            name := IniRead(configFile, section, "Name", "")
            if (name != "")
                apartmentNames.Push(name)
        }
    }
    if (apartmentNames.Length > 0) {
        sortedNames := SortArray(apartmentNames)
        for name in sortedNames
            AptSelectDropdown.Add([name])
        try {
            AptSelectDropdown.Choose(1)
        }
    }
}

SortArray(arr) {
    sortedArr := []
    for item in arr
        sortedArr.Push(item)
    loop sortedArr.Length - 1 {
        i := A_Index
        loop sortedArr.Length - i {
            j := A_Index
            if (StrCompare(sortedArr[j], sortedArr[j + 1]) > 0) {
                temp := sortedArr[j]
                sortedArr[j] := sortedArr[j + 1]
                sortedArr[j + 1] := temp
            }
        }
    }
    return sortedArr
}

SaveData(*) {
    global configFile, CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl, StatusBar
    if !FileExist(configFile)
        FileAppend("", configFile, "UTF-16")
    try {
        IniWrite(CleanCtrl.Value, configFile, "Expenses", "Clean")
        IniWrite(ElectricityCtrl.Value, configFile, "Expenses", "Electricity")
        IniWrite(WaterCtrl.Value, configFile, "Expenses", "Water")
        IniWrite(FireCtrl.Value, configFile, "Expenses", "Fire")
        IniWrite(GardenerCtrl.Value, configFile, "Expenses", "Gardener")
        IniWrite(OtherCtrl.Value, configFile, "Expenses", "Other")
        IniWrite(ElevatorCtrl.Value, configFile, "Expenses", "Elevator")
        IniWrite(HeatingCtrl.Value, configFile, "Expenses", "Heating")
        IniWrite(PrintingCtrl.Value, configFile, "Expenses", "Printing")
        IniWrite(ReserveCtrl.Value, configFile, "Expenses", "Reserve")
        StatusBar.SetText("✅ Τα δεδομένα αποθηκεύτηκαν επιτυχώς!")
        MsgBox("Τα δεδομένα αποθηκεύτηκαν επιτυχώς!", "Αποθήκευση", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά την αποθήκευση!")
        MsgBox("Σφάλμα κατά την αποθήκευση: " . e.Message, "Σφάλμα", "Icon!")
    }
}

LoadData(*) {
    global configFile, CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl, StatusBar
    if !FileExist(configFile) {
        MsgBox("Το αρχείο δεδομένων δεν βρέθηκε. Θα δημιουργηθεί αυτόματα.", "Πληροφορία", "Iconi")
        FileAppend("", configFile, "UTF-16")
        return
    }
    try {
        CleanCtrl.Value := IniRead(configFile, "Expenses", "Clean", "0.00")
        ElectricityCtrl.Value := IniRead(configFile, "Expenses", "Electricity", "0.00")
        WaterCtrl.Value := IniRead(configFile, "Expenses", "Water", "0.00")
        FireCtrl.Value := IniRead(configFile, "Expenses", "Fire", "0.00")
        GardenerCtrl.Value := IniRead(configFile, "Expenses", "Gardener", "0.00")
        OtherCtrl.Value := IniRead(configFile, "Expenses", "Other", "0.00")
        ElevatorCtrl.Value := IniRead(configFile, "Expenses", "Elevator", "0.00")
        HeatingCtrl.Value := IniRead(configFile, "Expenses", "Heating", "0.00")
        PrintingCtrl.Value := IniRead(configFile, "Expenses", "Printing", "0.00")
        ReserveCtrl.Value := IniRead(configFile, "Expenses", "Reserve", "0.00")
        CalculateTotalExpenses()
        StatusBar.SetText("✅ Τα δεδομένα φορτώθηκαν επιτυχώς!")
        MsgBox("Τα δεδομένα φορτώθηκαν επιτυχώς!", "Φόρτωση", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά τη φόρτωση!")
        MsgBox("Σφάλμα κατά τη φόρτωση: " . e.Message, "Σφάλμα", "Icon!")
    }
}

LoadApartments(*) {
    global configFile, AptsListView, StatusBar
    AptsListView.Delete()
    if !FileExist(configFile) {
        MsgBox("Το αρχείο δεδομένων δεν βρέθηκε. Θα δημιουργηθεί αυτόματα.", "Πληροφορία", "Iconi")
        FileAppend("", configFile, "UTF-16")
        return
    }
    sections := IniRead(configFile)
    if (sections = "ERROR") {
        MsgBox("Δεν βρέθηκαν διαμερίσματα. Προσθέστε πρώτα κάποια διαμερίσματα.", "Πληροφορία", "Iconi")
        return
    }
    apartmentCount := 0
    totalPayment := 0
    loop parse, sections, "`n" {
        section := Trim(A_LoopField)
        if (InStr(section, "Apartment_")) {
            apartmentCount++
            name := IniRead(configFile, section, "Name", "")
            owner := IniRead(configFile, section, "Owner", "")
            common := IniRead(configFile, section, "CommonPercent", "0")
            elevator := IniRead(configFile, section, "ElevatorPercent", "0")
            heating := IniRead(configFile, section, "HeatingPercent", "0")
            printing := IniRead(configFile, section, "PrintingPercent", "0")
            payment := CalculateApartmentPayment(name)
            totalPayment += payment
            if !IsNumber(payment)
                payment := 0
            AptsListView.Add("", name, owner, common, elevator, heating, printing, Format("{:.2f} €", payment))
        }
    }
    if (apartmentCount > 0) {
        StatusBar.SetText("✅ Φορτώθηκαν " . apartmentCount . " διαμερίσματα")
        MsgBox("Φορτώθηκαν " . apartmentCount . " διαμερίσματα.", "Πληροφορία", "Iconi")
    }
    totalCommon := 0
    totalElevator := 0
    totalHeating := 0
    loop apartmentCount {
        totalCommon += Integer(AptsListView.GetText(A_Index, 3))
        totalElevator += Integer(AptsListView.GetText(A_Index, 4))
        totalHeating += Integer(AptsListView.GetText(A_Index, 5))
    }
    if (apartmentCount > 0)
        AptsListView.Add("", "ΣΥΝΟΛΟ", "", totalCommon, totalElevator, totalHeating, "100.00", Format("{:.2f} €", totalPayment))
    AptsListView.ModifyCol(1, 120)
    AptsListView.ModifyCol(2, 180)
    AptsListView.ModifyCol(3, 100)
    AptsListView.ModifyCol(4, 100)
    AptsListView.ModifyCol(5, 110)
    AptsListView.ModifyCol(6, 100)
    AptsListView.ModifyCol(7, 120)
    UpdateApartmentDropdown()
}

AddApartment(*) {
    global configFile, AptNameCtrl, AptOwnerCtrl, AptCommonCtrl, AptElevatorCtrl, AptHeatingCtrl, StatusBar
    name := AptNameCtrl.Value
    owner := AptOwnerCtrl.Value
    common := AptCommonCtrl.Value
    elevator := AptElevatorCtrl.Value
    heating := AptHeatingCtrl.Value
    if (name = "") {
        MsgBox("Παρακαλώ εισάγετε όνομα διαμερίσματος!", "Προσοχή", "Icon!")
        return
    }
    if !FileExist(configFile)
        FileAppend("", configFile, "UTF-16")
    try {
        section := "Apartment_" . name
        IniWrite(name, configFile, section, "Name")
        IniWrite(owner, configFile, section, "Owner")
        IniWrite(common, configFile, section, "CommonPercent")
        IniWrite(elevator, configFile, section, "ElevatorPercent")
        IniWrite(heating, configFile, section, "HeatingPercent")
        UpdatePrintingPercentages()
        LoadApartments()
        AptNameCtrl.Value := ""
        AptOwnerCtrl.Value := ""
        AptCommonCtrl.Value := "0"
        AptElevatorCtrl.Value := "0"
        AptHeatingCtrl.Value := "0"
        StatusBar.SetText("✅ Το διαμέρισμα '" . name . "' προστέθηκε επιτυχώς!")
        MsgBox("Το διαμέρισμα προστέθηκε επιτυχώς!", "Επιτυχία", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά την προσθήκη!")
        MsgBox("Σφάλμα κατά την προσθήκη: " . e.Message, "Σφάλμα", "Icon!")
    }
}

SaveApartmentChanges(*) {
    global configFile, currentEditingApt, AptNameCtrl, AptOwnerCtrl, AptCommonCtrl, AptElevatorCtrl, AptHeatingCtrl, StatusBar
    if (currentEditingApt = "") {
        MsgBox("Δεν υπάρχει διαμέρισμα προς επεξεργασία!", "Προσοχή", "Icon!")
        return
    }
    name := AptNameCtrl.Value
    owner := AptOwnerCtrl.Value
    common := AptCommonCtrl.Value
    elevator := AptElevatorCtrl.Value
    heating := AptHeatingCtrl.Value
    if (name = "") {
        MsgBox("Παρακαλώ εισάγετε όνομα διαμερίσματος!", "Προσοχή", "Icon!")
        return
    }
    try {
        oldSection := "Apartment_" . currentEditingApt
        IniDelete(configFile, oldSection)
        section := "Apartment_" . name
        IniWrite(name, configFile, section, "Name")
        IniWrite(owner, configFile, section, "Owner")
        IniWrite(common, configFile, section, "CommonPercent")
        IniWrite(elevator, configFile, section, "ElevatorPercent")
        IniWrite(heating, configFile, section, "HeatingPercent")
        UpdatePrintingPercentages()
        LoadApartments()
        AptNameCtrl.Value := ""
        AptOwnerCtrl.Value := ""
        AptCommonCtrl.Value := "0"
        AptElevatorCtrl.Value := "0"
        AptHeatingCtrl.Value := "0"
        currentEditingApt := ""
        StatusBar.SetText("✅ Οι αλλαγές αποθηκεύτηκαν επιτυχώς!")
        MsgBox("Οι αλλαγές αποθηκεύτηκαν επιτυχώς!", "Επιτυχία", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά την αποθήκευση!")
        MsgBox("Σφάλμα κατά την αποθήκευση των αλλαγών: " . e.Message, "Σφάλμα", "Icon!")
    }
}

DeleteApartment(*) {
    global configFile, AptsListView, StatusBar
    selectedRow := AptsListView.GetNext()
    if (selectedRow = 0) {
        MsgBox("Παρακαλώ επιλέξτε ένα διαμέρισμα για διαγραφή!", "Προσοχή", "Icon!")
        return
    }
    aptName := AptsListView.GetText(selectedRow, 1)
    if (MsgBox("Είστε σίγουρος ότι θέλετε να διαγράψετε το διαμέρισμα '" . aptName . "'?", "Επιβεβαίωση", "YesNo Icon!") = "Yes") {
        try {
            section := "Apartment_" . aptName
            IniDelete(configFile, section)
            UpdatePrintingPercentages()
            LoadApartments()
            StatusBar.SetText("✅ Το διαμέρισμα '" . aptName . "' διαγράφτηκε!")
            MsgBox("Το διαμέρισμα διαγράφτηκε επιτυχώς!", "Επιτυχία", "Iconi")
        } catch as e {
            StatusBar.SetText("❌ Σφάλμα κατά τη διαγραφή!")
            MsgBox("Σφάλμα κατά τη διαγραφή: " . e.Message, "Σφάλμα", "Icon!")
        }
    }
}

EditApartment(*) {
    global configFile, AptsListView, currentEditingApt, AptNameCtrl, AptOwnerCtrl, AptCommonCtrl, AptElevatorCtrl, AptHeatingCtrl, StatusBar
    selectedRow := AptsListView.GetNext()
    if (selectedRow = 0) {
        MsgBox("Παρακαλώ επιλέξτε ένα διαμέρισμα για επεξεργασία!", "Προσοχή", "Icon!")
        return
    }
    aptName := AptsListView.GetText(selectedRow, 1)
    section := "Apartment_" . aptName
    try {
        AptNameCtrl.Value := aptName
        AptOwnerCtrl.Value := IniRead(configFile, section, "Owner", "")
        AptCommonCtrl.Value := IniRead(configFile, section, "CommonPercent", "0")
        AptElevatorCtrl.Value := IniRead(configFile, section, "ElevatorPercent", "0")
        AptHeatingCtrl.Value := IniRead(configFile, section, "HeatingPercent", "0")
        currentEditingApt := aptName
        StatusBar.SetText("✏️ Επεξεργασία διαμερίσματος: " . aptName)
        MsgBox("Τα δεδομένα του διαμερίσματος φορτώθηκαν για επεξεργασία!", "Επεξεργασία", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά τη φόρτωση!")
        MsgBox("Σφάλμα κατά τη φόρτωση: " . e.Message, "Σφάλμα", "Icon!")
    }
}
CalculateApartmentExpenses(*) {
    global configFile, AptSelectDropdown, StatusBar
    global AptCommonBaseCtrl, AptCommonPercentCtrl, AptCommonResultCtrl
    global AptElevatorBaseCtrl, AptElevatorPercentCtrl, AptElevatorResultCtrl
    global AptHeatingBaseCtrl, AptHeatingPercentCtrl, AptTotalHeatingPercentCtrl, AptHeatingResultCtrl
    global AptPrintingBaseCtrl, AptPrintingPercentCtrl, AptPrintingResultCtrl
    global AptReserveBaseCtrl, AptReservePercentCtrl, AptReserveResultCtrl, AptTotalResultCtrl
    global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl
    selectedApt := AptSelectDropdown.Text
    if (selectedApt = "")
        return 0
    totalCommon := CalculateTotalCommon()
    totalElevator := ConvertToNumber(ElevatorCtrl.Value)
    totalHeating := ConvertToNumber(HeatingCtrl.Value)
    totalPrinting := ConvertToNumber(PrintingCtrl.Value)
    totalReserve := ConvertToNumber(ReserveCtrl.Value)
    section := "Apartment_" . selectedApt
    commonPercent := ConvertToNumber(IniRead(configFile, section, "CommonPercent", "0"))
    elevatorPercent := ConvertToNumber(IniRead(configFile, section, "ElevatorPercent", "0"))
    heatingPercent := ConvertToNumber(IniRead(configFile, section, "HeatingPercent", "0"))
    printingPercent := ConvertToNumber(IniRead(configFile, section, "PrintingPercent", "0"))
    totalHeatingPercent := 0
    sections := IniRead(configFile)
    if (sections != "ERROR") {
        loop parse, sections, "`n" {
            sectionTemp := Trim(A_LoopField)
            if (InStr(sectionTemp, "Apartment_"))
                totalHeatingPercent += ConvertToNumber(IniRead(configFile, sectionTemp, "HeatingPercent", "0"))
        }
    }
    commonResult := (totalCommon * commonPercent) / 1000
    elevatorResult := (totalElevator * elevatorPercent) / 1000
    heatingResult := (totalHeatingPercent > 0) ? (totalHeating * heatingPercent) / totalHeatingPercent : 0
    printingResult := (totalPrinting * printingPercent) / 100
    reserveResult := (totalReserve * commonPercent) / 1000
    totalResult := commonResult + elevatorResult + heatingResult + printingResult + reserveResult
    AptCommonBaseCtrl.Value := Format("{:.2f} €", totalCommon)
    AptCommonPercentCtrl.Value := commonPercent
    AptCommonResultCtrl.Value := Format("{:.2f} €", commonResult)
    AptElevatorBaseCtrl.Value := Format("{:.2f} €", totalElevator)
    AptElevatorPercentCtrl.Value := elevatorPercent
    AptElevatorResultCtrl.Value := Format("{:.2f} €", elevatorResult)
    AptHeatingBaseCtrl.Value := Format("{:.2f} €", totalHeating)
    AptHeatingPercentCtrl.Value := heatingPercent
    AptTotalHeatingPercentCtrl.Value := totalHeatingPercent
    AptHeatingResultCtrl.Value := Format("{:.2f} €", heatingResult)
    AptPrintingBaseCtrl.Value := Format("{:.2f} €", totalPrinting)
    AptPrintingPercentCtrl.Value := Format("{:.2f}", printingPercent)
    AptPrintingResultCtrl.Value := Format("{:.2f} €", printingResult)
    AptReserveBaseCtrl.Value := Format("{:.2f} €", totalReserve)
    AptReservePercentCtrl.Value := commonPercent
    AptReserveResultCtrl.Value := Format("{:.2f} €", reserveResult)
    AptTotalResultCtrl.Value := Format("{:.2f} €", totalResult)
    StatusBar.SetText("✅ Υπολογισμός για: " . selectedApt . " | Σύνολο: " . Format("{:.2f} €", totalResult))
    return totalResult
}

CalculateTotalCommon() {
    global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    return ConvertToNumber(CleanCtrl.Value) + ConvertToNumber(ElectricityCtrl.Value) + 
           ConvertToNumber(WaterCtrl.Value) + ConvertToNumber(FireCtrl.Value) + 
           ConvertToNumber(GardenerCtrl.Value) + ConvertToNumber(OtherCtrl.Value)
}

CalculateTotalExpenses() {
    global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl
    global CommonTotalCtrl, ElevatorTotalCtrl, HeatingTotalCtrl, PrintingTotalCtrl, ReserveTotalCtrl, GrandTotalCtrl
    totalCommon := CalculateTotalCommon()
    totalElevator := ConvertToNumber(ElevatorCtrl.Value)
    totalHeating := ConvertToNumber(HeatingCtrl.Value)
    totalPrinting := ConvertToNumber(PrintingCtrl.Value)
    totalReserve := ConvertToNumber(ReserveCtrl.Value)
    grandTotal := totalCommon + totalElevator + totalHeating + totalPrinting + totalReserve
    CommonTotalCtrl.Value := Format("{:.2f} €", totalCommon)
    ElevatorTotalCtrl.Value := Format("{:.2f} €", totalElevator)
    HeatingTotalCtrl.Value := Format("{:.2f} €", totalHeating)
    PrintingTotalCtrl.Value := Format("{:.2f} €", totalPrinting)
    ReserveTotalCtrl.Value := Format("{:.2f} €", totalReserve)
    GrandTotalCtrl.Value := Format("{:.2f} €", grandTotal)
    return grandTotal
}

AutoCalculateExpenses(*) {
    CalculateTotalExpenses()
    UpdateApartmentPayments()
}

UpdateApartmentPayments() {
    global AptsListView, configFile
    rowCount := AptsListView.GetCount()
    if (rowCount <= 0)
        return
    loop rowCount - 1 {
        aptName := AptsListView.GetText(A_Index, 1)
        if (aptName != "ΣΥΝΟΛΟ") {
            payment := CalculateApartmentPayment(aptName)
            AptsListView.Modify(A_Index, "Col7", Format("{:.2f} €", payment))
        }
    }
    totalPayment := 0
    loop rowCount - 1 {
        aptName := AptsListView.GetText(A_Index, 1)
        if (aptName != "ΣΥΝΟΛΟ")
            totalPayment += CalculateApartmentPayment(aptName)
    }
    AptsListView.Modify(rowCount, "Col7", Format("{:.2f} €", totalPayment))
}

IsValidNumber(value) {
    if value = ""
        return true
    cleanValue := Trim(StrReplace(value, "€", ""))
    return RegExMatch(cleanValue, "^-?\d+([,.]\d+)?$")
}

ConvertToNumber(value) {
    if value = ""
        return 0
    cleanValue := Trim(StrReplace(value, "€", ""))
    cleanValue := StrReplace(cleanValue, ",", ".")
    return Number(cleanValue)
}

IsNumber(value) {
    return value is Number
}

AutoCalculateApartment(*) {
    CalculateApartmentExpenses()
}

ClearData(*) {
    global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl, StatusBar
    CleanCtrl.Value := "0.00"
    ElectricityCtrl.Value := "0.00"
    WaterCtrl.Value := "0.00"
    FireCtrl.Value := "0.00"
    GardenerCtrl.Value := "0.00"
    OtherCtrl.Value := "0.00"
    ElevatorCtrl.Value := "0.00"
    HeatingCtrl.Value := "0.00"
    PrintingCtrl.Value := "0.00"
    ReserveCtrl.Value := "0.00"
    CalculateTotalExpenses()
    StatusBar.SetText("🗑️ Τα δεδομένα καθαρίστηκαν!")
    MsgBox("Τα δεδομένα καθαρίστηκαν!", "Καθαρισμός", "Iconi")
}

SaveResultsToFile(*) {
    global configFile, AptSelectDropdown, StatusBar
    global AptCommonBaseCtrl, AptCommonPercentCtrl, AptCommonResultCtrl
    global AptElevatorBaseCtrl, AptElevatorPercentCtrl, AptElevatorResultCtrl
    global AptHeatingBaseCtrl, AptHeatingPercentCtrl, AptTotalHeatingPercentCtrl, AptHeatingResultCtrl
    global AptPrintingBaseCtrl, AptPrintingPercentCtrl, AptPrintingResultCtrl
    global AptReserveBaseCtrl, AptReservePercentCtrl, AptReserveResultCtrl, AptTotalResultCtrl
    selectedApt := AptSelectDropdown.Text
    if (selectedApt = "") {
        MsgBox("Παρακαλώ επιλέξτε ένα διαμέρισμα!", "Προσοχή", "Icon!")
        return
    }
    total := CalculateApartmentExpenses()
    fileName := "Αποτελέσματα_" . selectedApt . "_" . A_YYYY . A_MM . A_DD . ".txt"
    totalHeatingPercent := 0
    sections := IniRead(configFile)
    if (sections != "ERROR") {
        loop parse, sections, "`n" {
            sectionTemp := Trim(A_LoopField)
            if (InStr(sectionTemp, "Apartment_"))
                totalHeatingPercent += Integer(IniRead(configFile, sectionTemp, "HeatingPercent", "0"))
        }
    }
    content := "═══════════════════════════════════════════════════════════`n"
    content .= "          ΑΠΟΤΕΛΕΣΜΑΤΑ ΔΙΑΜΕΡΙΣΜΑΤΟΣ: " . selectedApt . "`n"
    content .= "═══════════════════════════════════════════════════════════`n`n"
    content .= "Ημερομηνία: " . A_DD . "/" . A_MM . "/" . A_YYYY . "`n`n"
    content .= "-----------------------------------------------------------`n"
    content .= "ΑΝΑΛΥΤΙΚΟΣ ΥΠΟΛΟΓΙΣΜΟΣ`n"
    content .= "-----------------------------------------------------------`n`n"
    content .= "ΚΟΙΝΟΧΡΗΣΤΑ:   " . Format("{:15}", AptCommonBaseCtrl.Value) . " × " 
    . Format("{:5}", AptCommonPercentCtrl.Value) . " / 1000 = " 
    . Format("{:10}", AptCommonResultCtrl.Value) . "`n"
    content .= "ΑΣΑΝΣΕΡ:       " . Format("{:15}", AptElevatorBaseCtrl.Value) . " × " 
    . Format("{:5}", AptElevatorPercentCtrl.Value) . " / 1000 = " 
    . Format("{:10}", AptElevatorResultCtrl.Value) . "`n"
    content .= "ΘΕΡΜΑΝΣΗ:      " . Format("{:15}", AptHeatingBaseCtrl.Value) . " × " 
    . Format("{:5}", AptHeatingPercentCtrl.Value) . " / " 
    . Format("{:5}", totalHeatingPercent) . " = " 
    . Format("{:10}", AptHeatingResultCtrl.Value) . "`n"
    content .= "ΕΚΔΟΣΗ:        " . Format("{:15}", AptPrintingBaseCtrl.Value) . " × " 
    . Format("{:5}", AptPrintingPercentCtrl.Value) . " / 100 = " 
    . Format("{:10}", AptPrintingResultCtrl.Value) . "`n"
    content .= "ΑΠΟΘΕΜΑΤΙΚΟ:   " . Format("{:15}", AptReserveBaseCtrl.Value) . " × " 
    . Format("{:5}", AptReservePercentCtrl.Value) . " / 1000 = " 
    . Format("{:10}", AptReserveResultCtrl.Value) . "`n"
    content .= "`n-----------------------------------------------------------`n"
    content .= "ΣΥΝΟΛΟ:                                        " . Format("{:10}", AptTotalResultCtrl.Value) . "`n"
    content .= "═══════════════════════════════════════════════════════════`n"
    try {
        FileAppend(content, fileName, "UTF-16")
        StatusBar.SetText("✅ Αποθηκεύτηκε: " . fileName)
        MsgBox("Τα αποτελέσματα αποθηκεύτηκαν στο αρχείο: " . fileName, "Αποθήκευση", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά την αποθήκευση!")
        MsgBox("Σφάλμα κατά την αποθήκευση: " . e.Message, "Σφάλμα", "Icon!")
    }
}

CalculateAndGoToTab3(*) {
    global TabCtrl, AptsListView, AptSelectDropdown, StatusBar
    selectedRow := AptsListView.GetNext()
    if (selectedRow = 0) {
        MsgBox("Παρακαλώ επιλέξτε ένα διαμέρισμα!", "Προσοχή", "Icon!")
        return
    }
    aptName := AptsListView.GetText(selectedRow, 1)
    try {
        AptSelectDropdown.Choose(aptName)
    }
    TabCtrl.Choose(3)
    CalculateApartmentExpenses()
    StatusBar.SetText("📈 Αναλυτικά για: " . aptName)
}

CalculateApartmentPayment(aptName) {
    global configFile, CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl
    totalCommon := CalculateTotalCommon()
    totalElevator := ConvertToNumber(ElevatorCtrl.Value)
    totalHeating := ConvertToNumber(HeatingCtrl.Value)
    totalPrinting := ConvertToNumber(PrintingCtrl.Value)
    totalReserve := ConvertToNumber(ReserveCtrl.Value)
    section := "Apartment_" . aptName
    commonPercent := ConvertToNumber(IniRead(configFile, section, "CommonPercent", "0"))
    elevatorPercent := ConvertToNumber(IniRead(configFile, section, "ElevatorPercent", "0"))
    heatingPercent := ConvertToNumber(IniRead(configFile, section, "HeatingPercent", "0"))
    printingPercent := ConvertToNumber(IniRead(configFile, section, "PrintingPercent", "0"))
    totalHeatingPercent := 0
    sections := IniRead(configFile)
    if (sections != "ERROR") {
        loop parse, sections, "`n" {
            sectionTemp := Trim(A_LoopField)
            if (InStr(sectionTemp, "Apartment_"))
                totalHeatingPercent += ConvertToNumber(IniRead(configFile, sectionTemp, "HeatingPercent", "0"))
        }
    }
    commonResult := (totalCommon * commonPercent) / 1000
    elevatorResult := (totalElevator * elevatorPercent) / 1000
    heatingResult := (totalHeatingPercent > 0) ? (totalHeating * heatingPercent) / totalHeatingPercent : 0
    printingResult := (totalPrinting * printingPercent) / 100
    reserveResult := (totalReserve * commonPercent) / 1000
    return commonResult + elevatorResult + heatingResult + printingResult + reserveResult
}

SaveTableToFile(*) {
    global AptsListView, StatusBar
    global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl
    fileName := "Πίνακας_Διαμερισμάτων_" . A_YYYY . A_MM . A_DD . ".txt"
    totalCommon := CalculateTotalCommon()
    totalElevator := ConvertToNumber(ElevatorCtrl.Value)
    totalHeating := ConvertToNumber(HeatingCtrl.Value)
    totalPrinting := ConvertToNumber(PrintingCtrl.Value)
    totalReserve := ConvertToNumber(ReserveCtrl.Value)
    grandTotal := totalCommon + totalElevator + totalHeating + totalPrinting + totalReserve
    content := "═══════════════════════════════════════════════════════════`n"
    content .= "           ΠΙΝΑΚΑΣ ΔΙΑΜΕΡΙΣΜΑΤΩΝ ΠΟΛΥΚΑΤΟΙΚΙΑΣ`n"
    content .= "═══════════════════════════════════════════════════════════`n`n"
    content .= "Ημερομηνία: " . A_DD . "/" . A_MM . "/" . A_YYYY . "`n`n"
    content .= "-----------------------------------------------------------`n"
    content .= "ΑΝΑΛΥΤΙΚΑ ΕΞΟΔΑ ΠΟΛΥΚΑΤΟΙΚΙΑΣ`n"
    content .= "-----------------------------------------------------------`n`n"
    content .= "ΚΑΘΑΡΙΟΤΗΤΑ:         " . Format("{:.2f} €", ConvertToNumber(CleanCtrl.Value)) . "`n"
    content .= "ΗΛΕΚΤΡΙΚΟ ΡΕΥΜΑ:     " . Format("{:.2f} €", ConvertToNumber(ElectricityCtrl.Value)) . "`n"
    content .= "ΝΕΡΟ:                " . Format("{:.2f} €", ConvertToNumber(WaterCtrl.Value)) . "`n"
    content .= "ΠΥΡΑΣΦΑΛΕΙΑ:         " . Format("{:.2f} €", ConvertToNumber(FireCtrl.Value)) . "`n"
    content .= "ΚΗΠΟΥΡΟΣ:            " . Format("{:.2f} €", ConvertToNumber(GardenerCtrl.Value)) . "`n"
    content .= "ΑΛΛΑ ΕΞΟΔΑ:          " . Format("{:.2f} €", ConvertToNumber(OtherCtrl.Value)) . "`n"
    content .= "                     " . "------------`n"
    content .= "ΣΥΝΟΛΟ ΚΟΙΝΟΧΡΗΣΤΩΝ: " . Format("{:.2f} €", totalCommon) . "`n`n"
    content .= "ΑΣΑΝΣΕΡ:             " . Format("{:.2f} €", totalElevator) . "`n"
    content .= "ΘΕΡΜАΝΣΗ:            " . Format("{:.2f} €", totalHeating) . "`n"
    content .= "ΕΚΔΟΣΗ:              " . Format("{:.2f} €", totalPrinting) . "`n"
    content .= "ΑΠΟΘΕΜΑΤΙΚΟ:         " . Format("{:.2f} €", totalReserve) . "`n"
    content .= "                     " . "============`n"
    content .= "ΓΕΝΙΚΟ ΣΥΝΟΛΟ:       " . Format("{:.2f} €", grandTotal) . "`n`n"
    content .= "═══════════════════════════════════════════════════════════`n"
    content .= "ΠΙΝΑΚΑΣ ΔΙΑΜΕΡΙΣΜΑΤΩΝ`n"
    content .= "═══════════════════════════════════════════════════════════`n`n"
    columnWidths := [15, 20, 15, 15, 15, 15, 15]
    headers := ""
    columnCount := AptsListView.GetCount("Column")
    loop columnCount {
        headerText := AptsListView.GetText(0, A_Index)
        headers .= Format("{:-" . columnWidths[A_Index] . "}", SubStr(headerText, 1, columnWidths[A_Index]))
    }
    content .= headers . "`n"
    content .= "-----------------------------------------------------------`n"
    rowCount := AptsListView.GetCount()
    loop rowCount {
        rowIndex := A_Index
        rowText := ""
        loop columnCount {
            cellText := AptsListView.GetText(rowIndex, A_Index)
            rowText .= Format("{:-" . columnWidths[A_Index] . "}", SubStr(cellText, 1, columnWidths[A_Index]))
        }
        content .= rowText . "`n"
        if (rowIndex = rowCount - 1)
            content .= "═══════════════════════════════════════════════════════════`n"
    }
    try {
        FileAppend(content, fileName, "UTF-16")
        StatusBar.SetText("✅ Πίνακας αποθηκεύτηκε: " . fileName)
        MsgBox("Ο πίνακας αποθηκεύτηκε επιτυχώς στο αρχείο: " . fileName, "Αποθήκευση Πίνακα", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά την αποθήκευση!")
        MsgBox("Σφάλμα κατά την αποθήκευση του πίνακα: " . e.Message, "Σφάλμα", "Icon!")
    }
}

SaveIniForPolyfund(*) {
    global configFile, AptsListView, StatusBar
    global CleanCtrl, ElectricityCtrl, WaterCtrl, FireCtrl, GardenerCtrl, OtherCtrl
    global ElevatorCtrl, HeatingCtrl, PrintingCtrl, ReserveCtrl
    fileName := "Expenses.ini"
    totalCommon := CalculateTotalCommon()
    totalElevator := ConvertToNumber(ElevatorCtrl.Value)
    totalHeating := ConvertToNumber(HeatingCtrl.Value)
    totalPrinting := ConvertToNumber(PrintingCtrl.Value)
    totalReserve := ConvertToNumber(ReserveCtrl.Value)
    grandTotal := totalCommon + totalElevator + totalHeating + totalPrinting + totalReserve
    if FileExist(fileName)
        FileDelete(fileName)
    try {
        IniWrite(A_DD . "/" . A_MM . "/" . A_YYYY, fileName, "General", "Date")
        IniWrite(Format("{:.2f}", grandTotal), fileName, "General", "TotalExpenses")
        IniWrite(Format("{:.2f}", ConvertToNumber(CleanCtrl.Value)), fileName, "Expenses", "Clean")
        IniWrite(Format("{:.2f}", ConvertToNumber(ElectricityCtrl.Value)), fileName, "Expenses", "Electricity")
        IniWrite(Format("{:.2f}", ConvertToNumber(WaterCtrl.Value)), fileName, "Expenses", "Water")
        IniWrite(Format("{:.2f}", ConvertToNumber(FireCtrl.Value)), fileName, "Expenses", "Fire")
        IniWrite(Format("{:.2f}", ConvertToNumber(GardenerCtrl.Value)), fileName, "Expenses", "Gardener")
        IniWrite(Format("{:.2f}", ConvertToNumber(OtherCtrl.Value)), fileName, "Expenses", "Other")
        IniWrite(Format("{:.2f}", totalCommon), fileName, "Expenses", "CommonTotal")
        IniWrite(Format("{:.2f}", totalElevator), fileName, "Expenses", "Elevator")
        IniWrite(Format("{:.2f}", totalHeating), fileName, "Expenses", "Heating")
        IniWrite(Format("{:.2f}", totalPrinting), fileName, "Expenses", "Printing")
        IniWrite(Format("{:.2f}", totalReserve), fileName, "Expenses", "Reserve")
        rowCount := AptsListView.GetCount()
        apartmentCount := 0
        loop rowCount {
            aptName := AptsListView.GetText(A_Index, 1)
            if (aptName = "ΣΥΝΟΛΟ")
                continue
            apartmentCount++
            owner := AptsListView.GetText(A_Index, 2)
            paymentText := AptsListView.GetText(A_Index, 7)
            payment := ConvertToNumber(paymentText)
            section := "Apartment_" . apartmentCount
            IniWrite(aptName, fileName, section, "Name")
            IniWrite(owner, fileName, section, "Owner")
            IniWrite(Format("{:.2f}", payment), fileName, section, "Payment")
            commonPercent := AptsListView.GetText(A_Index, 3)
            elevatorPercent := AptsListView.GetText(A_Index, 4)
            heatingPercent := AptsListView.GetText(A_Index, 5)
            printingPercent := AptsListView.GetText(A_Index, 6)
            IniWrite(commonPercent, fileName, section, "CommonPercent")
            IniWrite(elevatorPercent, fileName, section, "ElevatorPercent")
            IniWrite(heatingPercent, fileName, section, "HeatingPercent")
            IniWrite(printingPercent, fileName, section, "PrintingPercent")
        }
        IniWrite(apartmentCount, fileName, "General", "ApartmentCount")
        StatusBar.SetText("✅ Αρχείο Polyfund δημιουργήθηκε: " . fileName)
        MsgBox("Τα δεδομένα αποθηκεύτηκαν επιτυχώς στο αρχείο: " . fileName . "`n`nΠεριεχόμενα:`n- Γενικές δαπάνες`n- Αναλυτικές δαπάνες ανά κατηγορία`n- Στοιχεία διαμερισμάτων (όνομα, ιδιοκτήτης, ποσό πληρωμής)`n- Ποσοστά διαμερισμάτων", "Αποθήκευση για Polyfund", "Iconi")
    } catch as e {
        StatusBar.SetText("❌ Σφάλμα κατά την αποθήκευση!")
        MsgBox("Σφάλμα κατά την αποθήκευση: " . e.Message, "Σφάλμα", "Icon!")
    }
}

ShowInfo(*) {
    global StatusBar
    infoText := ""
    infoText .= "═══════════════════════════════════════════════════`n"
    infoText .= "                    Polycalc`n"
    infoText .= "═══════════════════════════════════════════════════`n`n"
    infoText .= "Έκδοση: v1.0`n"
    infoText .= "Δημιουργός: Tasos`n"
    infoText .= "Ημερομηνία Έκδοσης: 27/09/2025`n"
    infoText .= "Τελευταία Ενημέρωση: 22/10/2025`n`n"
    infoText .= "Email: maxiths1984@gmail.com`n`n"
    infoText .= "═══════════════════════════════════════════════════`n"
    infoText .= "© 2025 Όλα τα δικαιώματα διατηρούνται`n"
    infoText .= "═══════════════════════════════════════════════════"
    StatusBar.SetText("ℹ️ Πληροφορίες προγράμματος")
    MsgBox(infoText, "Πληροφορίες Προγράμματος", 64)
}

FilterNumericInput(Ctrl, Info) {
    text := Ctrl.Value
    if (text = "")
        return
    newText := ""
    allowedChars := "0123456789,."
    Loop Parse, text {
        if (InStr(allowedChars, A_LoopField))
            newText .= A_LoopField
    }
    if (newText != text)
        Ctrl.Value := newText
    commaCount := 0
    dotCount := 0
    Loop Parse, newText {
        if (A_LoopField = ",")
            commaCount++
        else if (A_LoopField = ".")
            dotCount++
    }
    if (commaCount > 1 || dotCount > 1 || (commaCount > 0 && dotCount > 0)) {
        firstCommaPos := InStr(newText, ",")
        firstDotPos := InStr(newText, ".")
        if (firstCommaPos > 0 && (firstDotPos = 0 || firstCommaPos < firstDotPos)) {
            finalText := SubStr(newText, 1, firstCommaPos) . StrReplace(SubStr(newText, firstCommaPos + 1), ",", "")
            finalText := StrReplace(finalText, ".", "")
        } else if (firstDotPos > 0) {
            finalText := SubStr(newText, 1, firstDotPos) . StrReplace(SubStr(newText, firstDotPos + 1), ".", "")
            finalText := StrReplace(finalText, ",", "")
        }
        Ctrl.Value := finalText
    }
}