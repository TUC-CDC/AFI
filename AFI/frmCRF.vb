
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class frmCRF
    Dim sql As New sqlControl
    Dim refer As Integer
    Dim dFeverOnsetdate As String = "NULL"
    Dim dLastJEDate As String = "NULL"
    Dim dAboardWhen As String = "NULL"
    Dim dBloodCollectDate1 As String = "NULL"
    Dim dBloodCollectDate2 As String = "NULL"
    Dim dBloodCollectTime1 As String = "NULL"
    Dim dBloodCollectTime2 As DateTime
    Dim dUrineCollectDate As String = "NULL"
    Dim dUrineCollectTime As DateTime
    Dim dCSFCollectDate As String = "NULL"
    Dim dCSFCollectTime As DateTime
    Dim dDischargeDate As String = "NULL"
    Dim isChecked As Boolean

    Private Sub setDateTimePickerBlank(ByVal dateTimePicker As DateTimePicker)

        dateTimePicker.Visible = True
        dateTimePicker.Format = DateTimePickerFormat.Custom
        dateTimePicker.CustomFormat = " "
        ' dateTimePicker.Enabled = False


    End Sub

    Function chkNumber(txtbox As TextBox) As String
        If Not String.IsNullOrEmpty(txtbox.Text.ToString.Trim) Then
            If IsNumeric(txtbox.Text.Trim) Then
                chkNumber = Double.Parse(txtbox.Text)
            Else
                MessageBox.Show(txtbox.Name & ": Number only")
                chkNumber = ""

            End If
        Else
                chkNumber = "Null"
        End If
    End Function
    Private Sub SetComboDate(cboD As ComboBox, cboM As ComboBox, cboY As ComboBox)
        For i = 1 To 31
            cboD.Items.Add(i)
        Next
        For i = 1 To 12
            cboM.Items.Add(MonthName(i))
        Next
        Dim y As Integer = DateAndTime.Now.Year

        For i = 2559 To y + 543
            cboY.Items.Add(i)
        Next
    End Sub
    Private Function SaveComboDate(cboD As ComboBox, cboM As ComboBox, cboY As ComboBox) As String

        If (String.IsNullOrEmpty(cboD.Text)) Or (String.IsNullOrEmpty(cboM.Text)) Or (String.IsNullOrEmpty(cboY.Text)) Then
            SaveComboDate = "NULL"

        Else

            SaveComboDate = cboY.Text - 543 & "/" & (cboM.Text) & "/" & cboD.Text

            SaveComboDate = "'" & Convert.ToDateTime(SaveComboDate) & "'"
        End If
    End Function

    Function chkCombo(combo As ComboBox) As String
        If Not String.IsNullOrEmpty(combo.SelectedValue) Then
            chkCombo = combo.SelectedValue
        Else
            chkCombo = "Null"
        End If
    End Function
    Private Sub frmCRF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If sql.HasConnection = True Then
            btnSave.Enabled = False
            SetComboDate(cboOnSetDay, cboOnSetMonth, cboOnSetYear)
            SetComboDate(cboAbroadWhenDay, cboAbroadWhenMonth, cboAboradWhenYear)
            SetComboDate(cboLastJEDay, cboLastJEMonth, cboLastJEYear)
            SetComboDate(cboBloodCollect1Day, cboBloodCollect1Month, cboBloodCollect1Year)
            SetComboDate(cboUrineCollectDay, cboUrineCollectMonth, cboUrineCollectYear)
            SetComboDate(cboCSFCollectDay, cboCSFCollectMonth, cboCSFCollectYear)
            SetComboDate(cboBloodCollect2Day, cboBloodCollect2Month, cboBloodCollect2Year)
            SetComboDate(cboDischargeDay, cboDischargeMonth, cboDischargeYear)

            sql.RunQuery("select * from lkuward where hospitalid =9", "Ward")
            If sql.SQLDataset.Tables.Count > 0 Then
                'cboWard.DataSource = sql.SQLDataset.Tables("Ward")
                'cboWard.ValueMember = "WardID"
                'cboWard.DisplayMember = "WardName"
            End If
            sql.RunQuery("select * from lkuAntibiotic", "Antibiotic1")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboAntibioticName1.DataSource = sql.SQLDataset.Tables("Antibiotic1")
                cboAntibioticName1.ValueMember = "AntibioticID"
                cboAntibioticName1.DisplayMember = "Antibiotic"
                cboAntibioticName1.Text = ""
            End If
            sql.RunQuery("select * from lkuAntibiotic", "Antibiotic2")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboAntibioticName2.DataSource = sql.SQLDataset.Tables("Antibiotic2")
                cboAntibioticName2.ValueMember = "AntibioticID"
                cboAntibioticName2.DisplayMember = "Antibiotic"
                cboAntibioticName2.Text = ""
            End If
            sql.RunQuery("select * from lkuAntibiotic", "Antibiotic3")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboAntibioticName3.DataSource = sql.SQLDataset.Tables("Antibiotic3")
                cboAntibioticName3.ValueMember = "AntibioticID"
                cboAntibioticName3.DisplayMember = "Antibiotic"
                cboAntibioticName3.Text = ""
            End If
            sql.RunQuery("select * from lkuHospital", "referHos")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboreferHos.DataSource = sql.SQLDataset.Tables("referHos")
                cboreferHos.ValueMember = "HospitalID"
                cboreferHos.DisplayMember = "HospitalName"

            End If
            sql.RunQuery("select * from lkuHospital", "TransferHosp")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboTransferHosp.DataSource = sql.SQLDataset.Tables("TransferHosp")
                cboTransferHosp.ValueMember = "HospitalID"
                cboTransferHosp.DisplayMember = "HospitalName"
                cboTransferHosp.Text = ""
            End If
            sql.RunQuery("select * from lkuUrineTaken", "UrineTaken")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboUrineTakenHow.DataSource = sql.SQLDataset.Tables("UrineTaken")
                cboUrineTakenHow.ValueMember = "MethodID"
                cboUrineTakenHow.DisplayMember = "Method"
            End If
            sql.RunQuery("select * from lkuOccupation", "Occupation1")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboOccupation1.DataSource = sql.SQLDataset.Tables("Occupation1")
                cboOccupation1.ValueMember = "OccupationID"
                cboOccupation1.DisplayMember = "OccupationTH"

            End If
            sql.RunQuery("select * from lkuOccupation", "Occupation2")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboOccupation2.DataSource = sql.SQLDataset.Tables("Occupation2")
                cboOccupation2.ValueMember = "OccupationID"
                cboOccupation2.DisplayMember = "OccupationTH"

            End If
            sql.RunQuery("select * from lkuGramStain", "GramStain")
            If sql.SQLDataset.Tables.Count > 0 Then
                cboCSF_GramStain.DataSource = sql.SQLDataset.Tables("GramStain")
                cboCSF_GramStain.ValueMember = "GramStainID"
                cboCSF_GramStain.DisplayMember = "GramStainResult"

            End If
        End If
    End Sub
    Private Function chkValidTime(msk As MaskedTextBox) As String
        With msk
            .TextMaskFormat = MaskFormat.IncludePromptAndLiterals
            chkValidTime = .Text

            .TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
            .HidePromptOnLeave = True
            If msk.Text <> "" Then
                chkValidTime = "'" & String.Format("{0:HH:mm}", chkValidTime) & "'"
            Else
                chkValidTime = "Null"
            End If
        End With

    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        Dim updCmd As String

        updCmd = "   Update  tblCase " &
                "SET StaffMember = '" & Trim(txtStaffMember.Text) & "'" &
                ",AdmitDiax = '" & Trim(txtAdmitDiax.Text) & "'" &
                ",Transferred = " & chk2Radio(rbTransferred1, rbTransferred2) &
                ",referHos =" & chkCombo(cboreferHos) &
                ",Weight = " & chkNumber(txtWeight) &
                ",Height = " & chkNumber(txtHeight) &
                ",Temp = " & chkNumber(txtTemp) &
                ",O2SatPercent =" & chkNumber(txtO2SatPercent) &
                ",O2SatRoomAir = " & chk3Radio(rbO2SatRoomAir1, rbO2SatRoomAir2, rbO2SatRoomAir3) &
                ",O2SatOxygenUsed =" & chkNumber(txtO2SatOxygenUsed) &
                ",RespRate = " & chkNumber(txtRespRate) &
                ",BPSys = " & chkNumber(txtBPSys) &
                ",BPDias =" & chkNumber(txtBPDias) &
                ",Pulse = " & chkNumber(txtPulse) &
                ",TourTest =  " & chk3Radio(rbTourTest1, rbTourTest2, rbTourTest3) &
                ",TourTestPetechiae = " & chkNumber(txtTourTestPetechiae) &
                ",Antibio72hrs = " & chk3Radio(rbAntibio72hrs1, rbAntibio72hrs2, rbAntibio72hrs3) &
                ",AntibioticName1 = " & chkCombo(cboAntibioticName1) &
                ",AntibioticNameOther1 = '" & Trim(txtAntibioticNameOther1.Text) & "'" &
                ",AntibioticIV_PO1 = " & chk2Radio(rbAntibioticIV_PO1_1, rbAntibioticIV_PO1_2) &
                ",AntibioticName2 =" & chkCombo(cboAntibioticName2) &
                ",AntibioticNameOther2 = '" & Trim(txtAntibioticNameOther2.Text) & "'" &
                ",AntibioticIV_PO2 = " & chk2Radio(rbAntibioticIV_PO2_1, rbAntibioticIV_PO2_2) &
                ",AntibioticName3 =" & chkCombo(cboAntibioticName3) &
                ",AntibioticNameOther3 = '" & Trim(txtAntibioticNameOther3.Text) & "'" &
                ",AntibioticIV_PO3 = " & chk2Radio(rbAntibioticIV_PO3_1, rbAntibioticIV_PO3_2) &
                ",FeverOnsetDate = " & SaveComboDate(cboOnSetDay, cboOnSetMonth, cboOnSetYear) &
                ",Diarrhea =" & chk3Radio(rbDiarrhea1, rbDiarrhea2, rbDiarrhea3) &
                ",Cough = " & chk3Radio(rbCough1, rbCough2, rbCough3) &
                ",RunnyNose =" & chk3Radio(rbRunnyNose1, rbRunnyNose2, rbRunnyNose3) &
                ",SoreThroat = " & chk3Radio(rbSoreThroat1, rbSoreThroat2, rbSoreThroat3) &
                ",ShortBreath = " & chk3Radio(rbShortBreath1, rbShortBreath2, rbShortBreath3) &
                ",SputumProduction = " & chk3Radio(rbSputumProduction1, rbSputumProduction2, rbSputumProduction3) &
                ",Hemoptysis =" & chk3Radio(rbHemoptysis1, rbHemoptysis2, rbHemoptysis3) &
                ",EyePain = " & chk3Radio(rbEyePain1, rbEyePain2, rbEyePain3) &
                ",RedEyes = " & chk3Radio(rbRedEyes1, rbRedEyes2, rbRedEyes3) &
                ",YellowEyes = " & chk3Radio(rbYellowEyes1, rbYellowEyes2, rbYellowEyes3) &
                ",Headache =" & chk3Radio(rbHeadache1, rbHeadache2, rbHeadache3) &
                ",Nausea = " & chk3Radio(rbNausea1, rbNausea2, rbNausea3) &
                ",NoseBleed = " & chk3Radio(rbNoseBleed1, rbNoseBleed2, rbNoseBleed3) &
                ",NeckStiff = " & chk3Radio(rbNeckStiff1, rbNeckStiff2, rbNeckStiff3) &
                ",Vomit = " & chk3Radio(rbVomit1, rbVomit2, rbVomit3) &
                ",BloodVomit =" & chk3Radio(rbBloodVomit1, rbBloodVomit2, rbBloodVomit3) &
                ",BloodStool = " & chk3Radio(rbBloodStool1, rbBloodStool2, rbBloodStool3) &
                ",BloodUrine = " & chk3Radio(rbBloodUrine1, rbBloodUrine2, rbBloodUrine3) &
                ",Dysuria =" & chk3Radio(rbDysuria1, rbDysuria2, rbDysuria3) &
                ",MusclePain = " & chk3Radio(rbMusclePain1, rbMusclePain2, rbMusclePain3) &
                ",CulfMusclePain =" & chk3Radio(rbCulfMusclePain1, rbCulfMusclePain2, rbCulfMusclePain3) &
                ",BonePain = " & chk3Radio(rbBonePain1, rbBonePain2, rbBonePain3) &
                ",ChestPain = " & chk3Radio(rbChestPain1, rbChestPain2, rbChestPain3) &
                ",BackPain =" & chk3Radio(rbBackPain1, rbBackPain2, rbBackPain3) &
                ",AbdominalPain = " & chk3Radio(rbAbdominalPain1, rbAbdominalPain2, rbAbdominalPain3) &
                ",JointPain =" & chk3Radio(rbJointPain1, rbJointPain2, rbJointPain3) &
                ",RedJoints = " & chk3Radio(rbRedJoints1, rbRedJoints2, rbRedJoints3) &
                ",NoAppetite =" & chk3Radio(rbNoAppetite1, rbNoAppetite2, rbNoAppetite3) &
                ",Tiredness = " & chk3Radio(rbTiredness1, rbTiredness2, rbTiredness3) &
                ",Seizures =" & chk3Radio(rbSeizures1, rbSeizures2, rbSeizures3) &
                ",Chills =" & chk3Radio(rbChills1, rbChills2, rbChills3) &
                ",Pale = " & chk3Radio(rbPale1, rbPale2, rbPale3) &
                ",Rash = " & chk3Radio(rbRash1, rbRash2, rbRash3) &
                ",Bruises = " & chk3Radio(rbBruises1, rbBruises2, rbBruises3) &
                ",ContactFlood =" & chk3Radio(rbContactFlood1, rbContactFlood2, rbContactFlood3) &
                ",ContactMud = " & chk3Radio(rbContactMud1, rbContactMud2, rbContactMud3) &
                ",ContactPond = " & chk3Radio(rbContactPond1, rbContactPond2, rbContactPond3) &
                ",WalkNoShoe =" & chk3Radio(rbWalkNoShoe1, rbWalkNoShoe2, rbWalkNoShoe3) &
                ",CutTree = " & chk3Radio(rbCutTree1, rbCutTree2, rbCutTree3) &
                ",VisitForest =" & chk3Radio(rbVisitForest1, rbVisitForest2, rbVisitForest3) &
                ",VisitRubberTree =" & chk3Radio(rbVisitRubberTree1, rbVisitRubberTree2, rbVisitRubberTree3) &
                ",CutSelf =" & chk3Radio(rbCutSelf1, rbCutSelf2, rbCutSelf3) &
                ",EatRawFish = " & chk3Radio(rbEatRawFish1, rbEatRawFish2, rbEatRawFish3) &
                ",EatRawPork = " & chk3Radio(rbEatRawPork1, rbEatRawPork2, rbEatRawPork3) &
                ",ContactAnimal =" & chk3Radio(rbContactAnimal1, rbContactAnimal2, rbContactAnimal3) &
                ",ConCaw = " & chk3Radio(rbConCaw1, rbConCaw2, rbConCaw3) &
                ",ConBuffalo = " & chk3Radio(rbConBuffalo1, rbConBuffalo2, rbConBuffalo3) &
                ", ConPig = " & chk3Radio(rbConPig1, rbConPig2, rbConPig3) &
                ",ConGoat =  " & chk3Radio(rbConGoat1, rbConGoat2, rbConGoat3) &
                ", ConSheep = " & chk3Radio(rbConSheep1, rbConSheep2, rbConSheep3) &
                ",ConChicken =  " & chk3Radio(rbConChicken1, rbConChicken2, rbConChicken3) &
                ", ConDuck = " & chk3Radio(rbConDuck1, rbConDuck2, rbConDuck3) &
                ",ConDog =  " & chk3Radio(rbConDog1, rbConDog2, rbConDog3) &
                ", ConCat = " & chk3Radio(rbConCat1, rbConCat2, rbConCat3) &
                ",ConRodent =  " & chk3Radio(rbConRodent1, rbConRodent2, rbConRodent3) &
                ", ConRat = " & chk3Radio(rbConRat1, rbConRat2, rbConRat3) &
                ",ConOthRodent =  " & chk3Radio(rbConOthRodent1, rbConOthRodent2, rbConOthRodent3) &
                ", ConOthRodentSpe = '" & txtConOthRodentSpe.Text.Trim & "'" &
                ",ConStrayAni =  " & chk3Radio(rbConStrayAni1, rbConStrayAni2, rbConStrayAni3) &
                ", ConStrayAniSpe = '" & txtConStrayAniSpe.Text.Trim & "'" &
                ",InsectBite =  " & chk3Radio(rbInsectBite1, rbInsectBite2, rbInsectBite3) &
                ", ConMosquitoes = " & chk3Radio(rbConMosquitoes1, rbConMosquitoes2, rbConMosquitoes3) &
                ",ConFebrileFamily =" & chk3Radio(rbConFebrileFamily1, rbConFebrileFamily2, rbConFebrileFamily3) &
                ",ConFebrileCowork =" & chk3Radio(rbConFebrileCowork1, rbConFebrileCowork2, rbConFebrileCowork3) &
                ",ConFebrileNeighbor =" & chk3Radio(rbConFebrileNeighbor1, rbConFebrileNeighbor2, rbConFebrileNeighbor3) &
                ",WentAbroad =" & chk3Radio(rbWentAboard1, rbWentAboard2, rbWentAboard3) &
                ",AbroadWhen =  " & SaveComboDate(cboAbroadWhenDay, cboAbroadWhenMonth, cboAboradWhenYear) &
                ",AbroadWhere = '" & txtAboradWhere.Text.Trim & "'" &
                ",VaccineJE  =" & chk3Radio(rbVaccineJE1, rbVaccineJE2, rbVaccineJE3) &
                ",LastJEDate =  " & dLastJEDate &
                ",LastJE =  " & chk4Radio(rbLastJE1, rbLastJE2, rbLastJE3, rbLastJE4) &
                ",JEInfoSource = " & chk2Radio(rbJEInfoSource1, rbJEInfoSource2) &
                ",Occupation1 = " & chkCombo(cboOccupation1) &
                ",OccupationOther1 = '" & txtOccupationOther1.Text.Trim & "'" &
                ",Occupation2 = " & chkCombo(cboOccupation2) &
                ",OccupationOther2 = '" & txtOccupationOther2.Text.Trim & "'" &
                ",Hematocrit = " & chkNumber(txtHematocrit) &
                ",Hematocrit_ND = " & chkCheckBox(chkHematocrit_ND) &
                ",Platelet = " & chkNumber(txtPlatelet) &
                ",Platelet_ND =" & chkCheckBox(chkPlatelet_ND) &
                ",WBCCount = " & chkNumber(txtWBCCount) &
                ",WBCCount_ND =" & chkCheckBox(chkWBCCount_ND) &
                ",NEUTROPHIL =  " & chkNumber(txtNEUTROPHIL) &
                ",NEUTROPHIL_ND =" & chkCheckBox(chkNEUTROPHIL_ND) &
                ",LYMPHOCYTE =  " & chkNumber(txtLYMPHOCYTE) &
                ",LYMPHOCYTE_ND =" & chkCheckBox(chkLYMPHOCYTE_ND) &
                ",MONOCYTE = " & chkNumber(txtMONOCYTE) &
                ",MONOCYTE_ND =" & chkCheckBox(chkMONOCYTE_ND) &
                ",EOSINOPHIL =  " & chkNumber(txtEOSINOPHIL) &
                ",EOSINOPHIL_ND = " & chkCheckBox(chkEOSINOPHIL_ND) &
                ",WBCNameOther =  '" & txtWBCNameOther.Text.Trim & "'" &
                ",WBCOtherResult =" & chkNumber(txtWBCOtherResult) &
                ",BUN = " & chkNumber(txtBUN) &
                ",BUN_ND = " & chkCheckBox(chkBUN_ND) &
                ",ALT = " & chkNumber(txtALT) &
                ",ALT_ND = " & chkCheckBox(chkALT_ND) &
                ",AST = " & chkNumber(txtAST) &
                ",AST_ND = " & chkCheckBox(chkAST_ND) &
                ",Creatinine = " & chkNumber(txtCreatinine) &
                ",Creatinine_ND = " & chkCheckBox(chkCreatinine_ND) &
                ",Albumin = " & chkNumber(txtAlbumin) &
                ",Albumin_ND = " & chkCheckBox(chkAlbumin_ND) &
                ",UrineTaken =" & chk2Radio(rbUrineTaken1, rbUrineTaken2) &
                ",UrineTAkenHow = " & chkCombo(cboUrineTakenHow) &
                ",RBCMore5 = " & chk2Radio(rbRBCMore51, rbRBCMore52) &
                ",WBCUrine = " & chkNumber(txtWBCUrine) &
                ",WBCUrine_ND = " & chkCheckBox(chkWBCUrine_ND) &
                ",CSF_WBC =" & chkNumber(txtCSF_WBC) &
                ",CSF_WBC_ND = " & chkCheckBox(chkCSF_WBC_ND) &
                ",CSF_RBC = " & chkNumber(txtCSF_RBC) &
                ",CSF_RBC_ND =" & chkCheckBox(chkCSF_RBC_ND) &
                ",CSF_Protein = " & chkNumber(txtCSF_Protein) &
                ",CSF_Protein_ND = " & chkCheckBox(chkCSF_Protein_ND) &
                ",CSF_Glucose = " & chkNumber(txtCSF_Glucose) &
                ",CSF_Glucose_ND = " & chkCheckBox(chkCSF_Glucose_ND) &
                ",CSF_GramStain =  " & chkCombo(cboCSF_GramStain) &
                ",CSF_GramStain_ND = " & chkCheckBox(chkCSF_GramStain_ND) &
                ",IsCultureBlood1 = " & chk2Radio(rbIsCultureBlood11, rbIsCultureBlood12) &
                ",BloodCollectDate1 =" & SaveComboDate(cboBloodCollect1Day, cboBloodCollect1Month, cboBloodCollect1Year) &
                ",BloodCollectTime1 =" & chkValidTime(mskBloodCollect1Time) &
                ",IsCultureUrine = " & chk2Radio(rbIsCultureUrine1, rbIsCultureUrine2) &
                ",UrineCollectDate =" & SaveComboDate(cboUrineCollectDay, cboUrineCollectMonth, cboUrineCollectYear) &
                ",UrineCollectTime =" & chkValidTime(mskUrineCollectTime) &
                ",IsCultureCSF = " & chk2Radio(rbIsCultureCSF1, rbIsCultureCSF2) &
                ",CSFCollectDate =" & SaveComboDate(cboCSFCollectDay, cboCSFCollectMonth, cboCSFCollectYear) &
                ",CSFCollectTime =" & chkValidTime(mskCSFCollectTime) &
                ",IsCultureBlood2 = " & chk2Radio(rbIsCultureBlood21, rbIsCultureBlood22) &
                ",BloodCollectDate2 =" & SaveComboDate(cboBloodCollect2Day, cboBloodCollect2Month, cboBloodCollect2Year) &
                ",BloodCollectTime2 =" & chkValidTime(mskBloodCollect2Time) &
                ",DengueIGM = " & chk2Radio(rbDengueIGM1, rbDengueIGM2) &
                ",DengueIGM_Ini = '" & txtDENGUEIGM_Ini.Text.Trim & "'" &
                ",DengueIGMResult = " & chk3Radio(rbDengueIGMResultPos, rbDengueIGMResultNeg, rbDengueIGMResultInd) &
                ",DengueNS = " & chk2Radio(rbDengueNS1, rbDengueNS2) &
                ",DengueNS_Ini = '" & txtDengueNS_Ini.Text.Trim & "'" &
                ",DengueNSResult = " & chk3Radio(rbDengueNSResultPos, rbDengueNSResultNeg, rbDengueNSResultInd) &
                ",PneuUri = " & chk2Radio(rbPneuUri1, rbPneuUri2) &
                ",PneuUri_Ini = '" & txtPneuUri_Ini.Text.Trim & "'" &
                ",PneuUriResult = " & chk3Radio(rbPneuUriResultPos, rbPneuUriResultNeg, rbPneuUriResultInd) &
                ",BPSerum = " & chk2Radio(rbBPSerum1, rbBPSerum2) &
                ",BPSerum_Ini = '" & txtBPSerum_Ini.Text.Trim & "'" &
                ",BPSerumResult = " & chk3Radio(rbBPSerumResultPos, rbBPSerumResultNeg, rbBPSerumResultInd) &
                ",BPUrine =" & chk2Radio(rbBPUrine1, rbBPUrine2) &
                ",BPUrine_Ini = '" & txtBPUrine_Ini.Text.Trim & "'" &
                ",BPUrineResult =  " & chk3Radio(rbBPUrineResultpos, rbBPUrineResultNeg, rbBPUrineResultInd) &
                ",BPSputum =" & chk2Radio(rbBPSputum1, rbBPSputum2) &
                ",BPSputum_Ini = '" & txtBPSputum_Ini.Text.Trim & "'" &
                ",BPSputumResult = " & chk3Radio(rbBPSputumResultPos, rbBPSputumResultNeg, rbBPSputumResultInd) &
                ",OtherRapidTest1 = '" & txtOtherRapidTest1.Text.Trim & "'" &
                ",DischargeDate =" & SaveComboDate(cboDischargeDay, cboDischargeMonth, cboDischargeYear) &
                ",Dengue = " & chk3Radio(rbDengue1, rbDengue2, rbDengue3) &
                ",Influenza = " & chk3Radio(rbInfluenza1, rbInfluenza2, rbInfluenza3) &
                ",Leptospirosis = " & chk3Radio(rbLeptospirosis1, rbLeptospirosis2, rbLeptospirosis3) &
                ",ScrubTyphus =" & chk3Radio(rbScrubTyphus1, rbScrubTyphus2, rbScrubTyphus3) &
                ",UpRespTract = " & chk3Radio(rbUpRespTract1, rbUpRespTract2, rbUpRespTract3) &
                ",UrineTract = " & chk3Radio(rbUrineTract1, rbUrineTract2, rbUrineTract3) &
                ",FeverUnknown =" & chk3Radio(rbFeverUnknown1, rbFeverUnknown2, rbFeverUnknown3) &
                ",InfectWound = " & chk3Radio(rbInfectWound1, rbInfectWound2, rbInfectWound3) &
                ",Typhoid =" & chk3Radio(rbTyphoid1, rbTyphoid2, rbTyphoid3) &
                ",Septicemia =" & chk3Radio(rbSepticemia1, rbSepticemia2, rbSepticemia3) &
                ",Gastroenteritis =" & chk3Radio(rbGastroenteritis1, rbGastroenteritis2, rbGastroenteritis3) &
                ",Melioidosis = " & chk3Radio(rbMelioidosis1, rbMelioidosis2, rbMelioidosis3) &
                ",Cellulitis = " & chk3Radio(rbCellulitis1, rbCellulitis2, rbCellulitis3) &
                ",Pneumonia = " & chk3Radio(rbPneumonia1, rbPneumonia2, rbPneumonia3) &
                ",Malaria = " & chk3Radio(rbMalaria1, rbMalaria2, rbMalaria3) &
                ",Chikungunya = " & chk3Radio(rbChikungunya1, rbChikungunya2, rbChikungunya3) &
                ",JE =" & chk3Radio(rbJE1, rbJE2, rbJE3) &
                ",Diabetes = " & chk2Radio(rbDiabetes1, rbDiabetes2) &
                ",Hypertension = " & chk2Radio(rbHypertension1, rbHypertension2) &
                ",HeartDisease = " & chk2Radio(rbHeartDisease1, rbHeartDisease2) &
                ",Asthma =" & chk2Radio(rbAsthma1, rbAsthma2) &
                ",CurSmoking =" & chk2Radio(rbCurSmoking1, rbCurSmoking2) &
                ",COPD = " & chk2Radio(rbCOPD1, rbCOPD2) &
                ",Cancer = " & chk2Radio(rbCancer1, rbCancer2) &
                ",CancerType ='" & txtCancerType.Text.Trim & "'" &
                ",Immunodeficiency = " & chk2Radio(rbImmunodeficiency1, rbImmunodeficiency2) &
                ",ImmunoSpe = '" & txtImmunoSpe.Text.Trim & "'" &
                ",KnownCase =" & chk2Radio(rbHIV1, rbHIV2) &
                ",HisTuber =" & chk2Radio(rbHisTuber1, rbHisTuber2) &
                ",ActTuber = " & chk2Radio(rbActTuber1, rbActTuber2) &
                ",Liver = " & chk2Radio(rbLiver1, rbLiver2) &
                ",Thyroid = " & chk2Radio(rbThyroid1, rbThyroid2) &
                ",Thalassemia = " & chk2Radio(rbThalassemia1, rbThalassemia2) &
                ",Anemia =" & chk2Radio(rbAnemia1, rbAnemia2) &
                ",ChroRenal = " & chk2Radio(rbChroRenal1, rbChroRenal2) &
                ",OtherCo1 = '" & txtOtherCo1.Text.Trim & "'" &
                ",OtherCo2 = '" & txtOtherCo2.Text.Trim & "'" &
                ",OtherCo3 = '" & txtOtherCo3.Text.Trim & "'" &
                ",Principal1 = '" & txtPrincipal1.Text.Trim & "'" &
                ",PrincipalICD10_1 ='" & txtPrincipalICD10_1.Text.Trim & "'" &
                ",Principal2 = '" & txtPrincipal2.Text.Trim & "'" &
                ",PrincipalICD10_2 ='" & txtPrincipalICD10_2.Text.Trim & "'" &
                ",Principal3 = '" & txtPrincipal3.Text.Trim & "'" &
                ",PrincipalICD10_3 ='" & txtPrincipalICD10_3.Text.Trim & "'" &
                ",Secondary1 = '" & txtSecondary1.Text.Trim & "'" &
                ",SecondaryICD10_1 = '" & txtSecondaryICD10_1.Text.Trim & "'" &
                ",Secondary2 = '" & txtSecondary2.Text.Trim & "'" &
                ",SecondaryICD10_2 = '" & txtSecondaryICD10_2.Text.Trim & "'" &
                ",Secondary3 = '" & txtSecondary3.Text.Trim & "'" &
                ",SecondaryICD10_3 = '" & txtSecondaryICD10_3.Text.Trim & "'" &
                ",Complication1 = '" & txtComplication1.Text.Trim & "'" &
                ",ComplicationICD10_1 = '" & txtComplicationICD10_1.Text.Trim & "'" &
                ",Complication2 = '" & txtComplication2.Text.Trim & "'" &
                ",ComplicationICD10_2 = '" & txtComplicationICD10_2.Text.Trim & "'" &
                ",Complication3 = '" & txtComplication3.Text.Trim & "'" &
                ",ComplicationICD10_3 = '" & txtComplicationICD10_3.Text.Trim & "'" &
                ",DischargeStatus = " & DischargeStatus() &
                ",DischargeType =  " & DischargeType() &
                ",TransferHospTo = " & chkCombo(cboTransferHosp) &
                ",TransferHospToOther = '" & txtTransferHospOth.Text.Trim & "'" &
                ",DischargeTypeOther = '" & txtDischargeTypeOther.Text.Trim & "'" &
                ",Intubation = " & chk3Radio(rbIntubation1, rbIntubation2, rbIntubation3) &
                ",_LastEditDate = '" & Date.Now & "'" &
                 " where afiid = '" & (txtAFIID.Text.Trim) & "'"

        sql.DataUpdate(updCmd)


    End Sub
    Private Function DischargeStatus() As Integer
        If rbDischargeStatus1.Checked = True Then
            DischargeStatus = 1
        ElseIf rbDischargeStatus2.Checked = True Then
            DischargeStatus = 2
        ElseIf rbDischargeStatus3.Checked = True Then
            DischargeStatus = 3
        ElseIf rbDischargeStatus4.Checked = True Then
            DischargeStatus = 4
        ElseIf rbDischargeStatus5.Checked = True Then
            DischargeStatus = 5
        Else
            DischargeStatus = 0
        End If
    End Function
    Private Function DischargeType() As Integer
        If rbDischargeType1.Checked = True Then
            DischargeType = 1
        ElseIf rbDischargeType2.Checked = True Then
            DischargeType = 2
        ElseIf rbDischargeType3.Checked = True Then
            DischargeType = 3
        ElseIf rbDischargeType4.Checked = True Then
            DischargeType = 4
        ElseIf rbDischargeType5.Checked = True Then
            DischargeType = 5
        ElseIf rbDischargeType6.Checked = True Then
            DischargeType = 6
        Else
            DischargeType = 0
        End If
    End Function
    Private Function chk2Radio(rb1 As RadioButton, rb2 As RadioButton) As Integer
        chk2Radio = 0
        If rb1.Checked = True Then
            chk2Radio = 1
        ElseIf rb2.Checked = True Then
            chk2Radio = 2

        End If
    End Function

    'Private Function chkDate(dt As DateTime) As String

    '    dt = DateTimePicker1.Value.Date ' use only the date portion '
    '    Dim tim As DateTime
    '    If DateTime.TryParse(TextBox1.Text, tim) Then
    '        dt = dt + tim.TimeOfDay ' use only the time portion '
    '    Else
    '        MsgBox("Could not understand the value in the time box. Using midnight.")
    '    End If

    'End Function
    Private Function chk3Radio(rb1 As RadioButton, rb2 As RadioButton, rb3 As RadioButton) As Integer
        chk3Radio = 0
        If rb1.Checked = True Then
            chk3Radio = 1
        ElseIf rb2.Checked = True Then
            chk3Radio = 2
        ElseIf rb3.Checked = True Then
            chk3Radio = 3
        Else
            chk3Radio = 0
        End If
    End Function
    Private Function chk4Radio(rb1 As RadioButton, rb2 As RadioButton, rb3 As RadioButton, rb4 As RadioButton) As Integer
        chk4Radio = 0
        If rb1.Checked = True Then
            chk4Radio = 1
        ElseIf rb2.Checked = True Then
            chk4Radio = 2
        ElseIf rb3.Checked = True Then
            chk4Radio = 3
        ElseIf rb4.Checked = True Then
            chk4Radio = 4
        Else
            chk4Radio = 0
        End If
    End Function
    Private Function chk5Radio(rb1 As RadioButton, rb2 As RadioButton, rb3 As RadioButton, rb4 As RadioButton, rb5 As RadioButton) As Integer
        chk5Radio = 0
        If rb1.Checked = True Then
            chk5Radio = 1
        ElseIf rb2.Checked = True Then
            chk5Radio = 2
        ElseIf rb3.Checked = True Then
            chk5Radio = 3
        ElseIf rb4.Checked = True Then
            chk5Radio = 4
        ElseIf rb5.Checked = True Then
            chk5Radio = 5
        Else
            chk5Radio = 0
        End If
    End Function
    Private Function chk6Radio(rb1 As RadioButton, rb2 As RadioButton, rb3 As RadioButton, rb4 As RadioButton, rb5 As RadioButton, rb6 As RadioButton) As Integer
        chk6Radio = 0
        If rb1.Checked = True Then
            chk6Radio = 1
        ElseIf rb2.Checked = True Then
            chk6Radio = 2
        ElseIf rb3.Checked = True Then
            chk6Radio = 3
        ElseIf rb4.Checked = True Then
            chk6Radio = 4
        ElseIf rb5.Checked = True Then
            chk6Radio = 5
        ElseIf rb6.Checked = True Then
            chk6Radio = 6
        Else
            chk6Radio = 0
        End If
    End Function
    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        Dim id As String

        id = Trim(txtAFIID.Text)


        If txtAFIID.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Add New Patient!")

        Else
            If sql.HasConnection = True Then
                resetAllControls(Me)
                sql.FindQuery("Select * from tblCase where AFIID='" & id & "'", id)
                txtAFIID.Text = id
            End If
            btnSave.Enabled = True
        End If


    End Sub
    Public Sub resetAllControls(ByVal container As Control)

        For Each ctrl As Control In container.Controls

            If TypeOf ctrl Is RadioButton Then

                DirectCast(ctrl, RadioButton).Checked = False

            End If

            If TypeOf ctrl Is ComboBox Then
                DirectCast(ctrl, ComboBox).SelectedValue = 0
            End If

            If TypeOf ctrl Is TextBox Then
                DirectCast(ctrl, TextBox).Clear()
            End If
            If TypeOf ctrl Is CheckBox Then
                DirectCast(ctrl, CheckBox).Checked = False
            End If
            If ctrl.Controls.Count > 0 Then

                resetAllControls(ctrl)

            End If

        Next

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    'Private Sub dtFeverOnsetDate_ValueChanged(sender As Object, e As EventArgs)

    '    'dtFeverOnsetDate.Format = DateTimePickerFormat.Short
    'End Sub

    'Private Sub chkOnsetDateBlank_CheckedChanged(sender As Object, e As EventArgs)
    '    If chkDeleteOnsetDate.Checked = True Then
    '        setDateTimePickerBlank(dtFeverOnsetDate)
    '        dtFeverOnsetDate.Enabled = False
    '    Else
    '        dtFeverOnsetDate.Enabled = True
    '        dtFeverOnsetDate.Format = DateTimePickerFormat.Short
    '    End If
    'End Sub

    Private Sub txtAFIID_KeyPress(sender As Object, e As EventArgs) Handles txtAFIID.KeyPress

        Dim tmp As System.Windows.Forms.KeyPressEventArgs = e
        If tmp.KeyChar = ChrW(Keys.Enter) Then
            MessageBox.Show("Enter key")
            Me.btnFind_Click(txtAFIID, e)
        Else
            ' MessageBox.Show(tmp.KeyChar)
        End If


    End Sub


    Private Sub rbWentAboard1_CheckedChanged(sender As Object, e As EventArgs) Handles rbWentAboard1.CheckedChanged, rbWentAboard2.CheckedChanged, rbWentAboard3.CheckedChanged
        If rbWentAboard1.Checked = True Then
            cboAbroadWhenDay.Enabled = True
            cboAbroadWhenMonth.Enabled = True
            cboAboradWhenYear.Enabled = True

        Else
            cboAbroadWhenDay.Enabled = False
            cboAbroadWhenMonth.Enabled = False
            cboAboradWhenYear.Enabled = False
        End If
    End Sub

    Private Sub rbVaccineJE1_CheckedChanged(sender As Object, e As EventArgs) Handles rbVaccineJE1.CheckedChanged
        If rbVaccineJE1.Checked = True Then
            cboLastJEDay.Enabled = True
            cboLastJEMonth.Enabled = True
            cboLastJEYear.Enabled = True
        Else
            cboLastJEDay.Enabled = False
            cboLastJEMonth.Enabled = False
            cboLastJEYear.Enabled = False
        End If
    End Sub



    Private Function chkCheckBox(chkbox As CheckBox) As Integer
        If chkbox.Checked = True Then
            chkCheckBox = 1
        Else
            chkCheckBox = 0
        End If

    End Function


    Private Sub gbRefer_Enter(sender As Object, e As EventArgs) Handles gbRefer.DoubleClick

        resetAllControls(gbRefer)
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter
        resetAllControls(Me.GroupBox2)
    End Sub

    Private Sub rbIsCultureBlood11_CheckedChanged(sender As Object, e As EventArgs) Handles rbIsCultureBlood11.CheckedChanged
        culture(rbIsCultureBlood11, cboBloodCollect1Day, cboBloodCollect1Month, cboBloodCollect1Year, mskBloodCollect1Time)
    End Sub

    Private Sub rbIsCultureUrine1_CheckedChanged(sender As Object, e As EventArgs) Handles rbIsCultureUrine1.CheckedChanged
        culture(rbIsCultureUrine1, cboUrineCollectDay, cboUrineCollectMonth, cboUrineCollectYear, mskUrineCollectTime)
    End Sub

    Private Sub culture(rb1 As RadioButton, cboD As ComboBox, cboM As ComboBox, cboY As ComboBox, mskTime As MaskedTextBox)
        If rb1.Checked = True Then
            cboD.Enabled = True
            cboM.Enabled = True
            cboY.Enabled = True
            mskTime.Enabled = True
        Else
            cboD.Enabled = False
            cboM.Enabled = False
            cboY.Enabled = False
            mskTime.Enabled = False
            cboD.SelectedText = ""
            cboM.SelectedText = ""
            cboY.SelectedText = ""
            mskTime.Text = ""

        End If
    End Sub

    Private Sub rbIsCultureBlood12_CheckedChanged(sender As Object, e As EventArgs) Handles rbIsCultureBlood12.CheckedChanged
        culture(rbIsCultureBlood12, cboBloodCollect2Day, cboBloodCollect2Month, cboBloodCollect2Year, mskBloodCollect2Time)
    End Sub


    Private Sub txtAFIID_TextChanged(sender As Object, e As EventArgs) Handles txtAFIID.TextChanged
        btnSave.Enabled = False
    End Sub
    Public Sub AddData2Radio(ByVal field As String, ByVal r1 As RadioButton, ByVal r2 As RadioButton)
        With sql.SQLDataset.Tables("aid")


            If IsDBNull(.Rows(0)(field)) Then
                r1.Checked = False
                r2.Checked = False
            Else

                If .Rows(0)(field) = 1 Then
                    r1.Checked = True
                ElseIf .Rows(0)(field) = 2 Then
                    r2.Checked = True
                End If
            End If
        End With
    End Sub

    Public Sub AddData3Radio(ByVal field As String, ByVal r1 As RadioButton, ByVal r2 As RadioButton, ByVal r3 As RadioButton)
        With sql.SQLDataset.Tables("aid")
            If IsDBNull(.Rows(0)(field)) Then
                r1.Checked = False
                r2.Checked = False
                r3.Checked = False
            Else
                If .Rows(0)(field) = 1 Then
                    r1.Checked = True
                ElseIf .Rows(0)(field) = 2 Then
                    r2.Checked = True
                ElseIf .Rows(0)(field) = 3 Then
                    r3.Checked = True
                End If

            End If
        End With
    End Sub
    Public Sub AddData4Radio(ByVal field As String, ByVal r1 As RadioButton, ByVal r2 As RadioButton, ByVal r3 As RadioButton, ByVal r4 As RadioButton)

        With sql.SQLDataset.Tables("aid")
            If IsDBNull(.Rows(0)(field)) Then
                r1.Checked = False
                r2.Checked = False
                r3.Checked = False
                r4.Checked = False
            Else
                If .Rows(0)(field) = 1 Then
                    r1.Checked = True
                ElseIf .Rows(0)(field) = 2 Then
                    r2.Checked = True
                ElseIf .Rows(0)(field) = 3 Then
                    r3.Checked = True

                ElseIf .Rows(0)(field) = 4 Then
                    r4.Checked = True

                End If
            End If
        End With
    End Sub
    Public Sub AddData5Radio(ByVal field As String, ByVal r1 As RadioButton, ByVal r2 As RadioButton, ByVal r3 As RadioButton, ByVal r4 As RadioButton, ByVal r5 As RadioButton)
        With sql.SQLDataset.Tables("aid")
            If IsDBNull(.Rows(0)(field)) Then
                r1.Checked = False
                r2.Checked = False
                r3.Checked = False
                r4.Checked = False
                r5.Checked = False
            Else
                If .Rows(0)(field) = 1 Then
                    r1.Checked = True
                ElseIf .Rows(0)(field) = 2 Then
                    r2.Checked = True
                ElseIf .Rows(0)(field) = 3 Then
                    r3.Checked = True

                ElseIf .Rows(0)(field) = 4 Then
                    r4.Checked = True
                ElseIf .Rows(0)(field) = 5 Then
                    r5.Checked = True
                End If
            End If
        End With
    End Sub
    Public Sub AddData6Radio(ByVal field As String, ByVal r1 As RadioButton, ByVal r2 As RadioButton, ByVal r3 As RadioButton, ByVal r4 As RadioButton, ByVal r5 As RadioButton, ByVal r6 As RadioButton)
        With sql.SQLDataset.Tables("aid")
            If IsDBNull(.Rows(0)(field)) Then
                r1.Checked = False
                r2.Checked = False
                r3.Checked = False
                r4.Checked = False
                r5.Checked = False
                r6.Checked = False
            Else
                If .Rows(0)(field) = 1 Then
                    r1.Checked = True
                ElseIf .Rows(0)(field) = 2 Then
                    r2.Checked = True
                ElseIf .Rows(0)(field) = 3 Then
                    r3.Checked = True

                ElseIf .Rows(0)(field) = 4 Then
                    r4.Checked = True
                ElseIf .Rows(0)(field) = 5 Then
                    r5.Checked = True
                ElseIf .Rows(0)(field) = 6 Then
                    r6.Checked = True
                End If
            End If
        End With
    End Sub
    Public Sub FillData()
        With sql.SQLDataset.Tables("aid")
            txtStaffMember.Text = .Rows(0)("StaffMember") & ""
            txtAdmitDiax.Text = .Rows(0)("AdmitDiax") & ""
            AddData2Radio("Transferred", rbTransferred1, rbTransferred2)
            cboreferHos.SelectedValue = IIf(.Rows(0)("referHos") Is Nothing, 0, .Rows(0)("referHos"))
            txtWeight.Text = IIf(IsDBNull(.Rows(0)("Weight")), "", .Rows(0)("Weight"))
            txtHeight.Text = .Rows(0)("Height") & ""
            txtTemp.Text = .Rows(0)("Temp") & ""
            txtO2SatPercent.Text = .Rows(0)("O2SatPercent") & ""
            AddData3Radio("O2SatRoomAir", rbO2SatRoomAir1, rbO2SatRoomAir2, rbO2SatRoomAir3)
            txtO2SatOxygenUsed.Text = .Rows(0)("O2SatOxygenUsed") & ""
            txtRespRate.Text = .Rows(0)("RespRate") & ""
            txtBPSys.Text = .Rows(0)("BPSys") & ""
            txtBPDias.Text = .Rows(0)("BPDias") & ""
            txtPulse.Text = .Rows(0)("Pulse") & ""
            AddData3Radio("TourTest", rbTourTest1, rbTourTest2, rbTourTest3)
            txtTourTestPetechiae.Text = .Rows(0)("TourTestPetechiae") & ""
            AddData3Radio("Antibio72hrs", rbAntibio72hrs1, rbAntibio72hrs2, rbAntibio72hrs3)
            cboAntibioticName1.SelectedValue = IIf(.Rows(0)("AntibioticName1") Is Nothing, 0, .Rows(0)("AntibioticName1"))
            txtAntibioticNameOther1.Text = .Rows(0)("AntibioticNameOther1") & ""
            AddData2Radio("AntibioticIV_PO1", rbAntibioticIV_PO1_1, rbAntibioticIV_PO1_2)
            cboAntibioticName2.SelectedValue = IIf(.Rows(0)("AntibioticName2") Is Nothing, 0, .Rows(0)("AntibioticName2"))
            txtAntibioticNameOther2.Text = .Rows(0)("AntibioticNameOther2") & ""
            AddData2Radio("AntibioticIV_PO2", rbAntibioticIV_PO2_1, rbAntibioticIV_PO2_2)
            cboAntibioticName3.SelectedValue = IIf(.Rows(0)("AntibioticName3") Is Nothing, 0, .Rows(0)("AntibioticName3"))
            txtAntibioticNameOther3.Text = .Rows(0)("AntibioticNameOther3") & ""
            AddData2Radio("AntibioticIV_PO3", rbAntibioticIV_PO3_1, rbAntibioticIV_PO3_2)

            If IsDBNull(.Rows(0)("FeverOnsetDate")) Then
                    SetComboDate(cboOnSetDay, cboOnSetMonth, cboOnSetYear)
                Else
                    cboOnSetDay.Text = .Rows(0)("FeverOnsetDate").day
                    cboOnSetMonth.Text = MonthName(.Rows(0)("FeverOnsetDate").month)

                    cboOnSetYear.Text = .Rows(0)("FeverOnsetDate").year + 543
                End If

            'DateTime.TryParse(.Rows(0)("AbroadWhen"), dtAboardWhen.Value)


            AddData3Radio("Diarrhea", rbDiarrhea1, rbDiarrhea2, rbDiarrhea3)
            AddData3Radio("Cough", rbCough1, rbCough2, rbCough3)
            AddData3Radio("RunnyNose", rbRunnyNose1, rbRunnyNose2, rbRunnyNose3)
            AddData3Radio("SoreThroat", rbSoreThroat1, rbSoreThroat2, rbSoreThroat3)
            AddData3Radio("ShortBreath", rbShortBreath1, rbShortBreath2, rbShortBreath3)
            AddData3Radio("SputumProduction", rbSputumProduction1, rbSputumProduction2, rbSputumProduction3)
            AddData3Radio("Hemoptysis", rbHemoptysis1, rbHemoptysis2, rbHemoptysis3)
            AddData3Radio("EyePain", rbEyePain1, rbEyePain2, rbEyePain3)
            AddData3Radio("RedEyes", rbRedEyes1, rbRedEyes2, rbRedEyes3)
            AddData3Radio("YellowEyes", rbYellowEyes1, rbYellowEyes2, rbYellowEyes3)
            AddData3Radio("Headache", rbHeadache1, rbHeadache2, rbHeadache3)
            AddData3Radio("Nausea", rbNausea1, rbNausea2, rbNausea3)
            AddData3Radio("NoseBleed", rbNoseBleed1, rbNoseBleed2, rbNoseBleed3)
            AddData3Radio("NeckStiff", rbNeckStiff1, rbNeckStiff2, rbNeckStiff3)
            AddData3Radio("Vomit", rbVomit1, rbVomit2, rbVomit3)
            AddData3Radio("BloodVomit", rbBloodVomit1, rbBloodVomit2, rbBloodVomit3)
            AddData3Radio("BloodStool", rbBloodStool1, rbBloodStool2, rbBloodStool3)
            AddData3Radio("BloodUrine", rbBloodUrine1, rbBloodUrine2, rbBloodUrine3)
            AddData3Radio("Dysuria", rbDysuria1, rbDysuria2, rbDysuria3)
            AddData3Radio("MusclePain", rbMusclePain1, rbMusclePain2, rbMusclePain3)
            AddData3Radio("CulfMusclePain", rbCulfMusclePain1, rbCulfMusclePain2, rbCulfMusclePain3)
            AddData3Radio("BonePain", rbBonePain1, rbBonePain2, rbBonePain3)
            AddData3Radio("ChestPain", rbChestPain1, rbChestPain2, rbChestPain3)
            AddData3Radio("BackPain", rbBackPain1, rbBackPain2, rbBackPain3)
            AddData3Radio("AbdominalPain", rbAbdominalPain1, rbAbdominalPain2, rbAbdominalPain3)
            AddData3Radio("JointPain", rbJointPain1, rbJointPain2, rbJointPain3)
            AddData3Radio("RedJoints", rbRedJoints1, rbRedJoints2, rbRedJoints3)
            AddData3Radio("NoAppetite", rbNoAppetite1, rbNoAppetite2, rbNoAppetite3)
            AddData3Radio("Tiredness", rbTiredness1, rbTiredness2, rbTiredness3)
            AddData3Radio("Seizures", rbSeizures1, rbSeizures2, rbSeizures3)
            AddData3Radio("Chills", rbChills1, rbChills2, rbChills3)
            AddData3Radio("Pale", rbPale1, rbPale2, rbPale3)
            AddData3Radio("Rash", rbRash1, rbRash2, rbRash3)
            AddData3Radio("Bruises", rbBruises1, rbBruises2, rbBruises3)
            AddData3Radio("ContactFlood", rbContactFlood1, rbContactFlood2, rbContactFlood3)
            AddData3Radio("ContactMud", rbContactMud1, rbContactMud2, rbContactMud3)
            AddData3Radio("ContactPond", rbContactPond1, rbContactPond2, rbContactPond3)
            AddData3Radio("WalkNoShoe", rbWalkNoShoe1, rbWalkNoShoe2, rbWalkNoShoe3)
            AddData3Radio("CutTree", rbCutTree1, rbCutTree2, rbCutTree3)
            AddData3Radio("VisitForest", rbVisitForest1, rbVisitForest2, rbVisitForest3)
            AddData3Radio("VisitRubberTree", rbVisitRubberTree1, rbVisitRubberTree2, rbVisitRubberTree3)
            AddData3Radio("CutSelf", rbCutSelf1, rbCutSelf2, rbCutSelf3)
            AddData3Radio("EatRawFish", rbEatRawFish1, rbEatRawFish2, rbEatRawFish3)
            AddData3Radio("EatRawPork", rbEatRawPork1, rbEatRawPork2, rbEatRawPork3)
            AddData3Radio("ContactAnimal", rbContactAnimal1, rbContactAnimal2, rbContactAnimal3)
            AddData3Radio("ConCaw", rbConCaw1, rbConCaw2, rbConCaw3)
            AddData3Radio("ConBuffalo", rbConBuffalo1, rbConBuffalo2, rbConBuffalo3)
            AddData3Radio("ConPig", rbConPig1, rbConPig2, rbConPig3)
            AddData3Radio("ConGoat", rbConGoat1, rbConGoat2, rbConGoat3)
            AddData3Radio("ConSheep", rbConSheep1, rbConSheep2, rbConSheep3)
            AddData3Radio("ConChicken", rbConChicken1, rbConChicken2, rbConChicken3)
            AddData3Radio("ConDuck", rbConDuck1, rbConDuck2, rbConDuck3)
            AddData3Radio("ConDog", rbConDog1, rbConDog2, rbConDog3)
            AddData3Radio("ConCat", rbConCat1, rbConCat2, rbConCat3)
            AddData3Radio("ConRodent", rbConRodent1, rbConRodent2, rbConRodent3)
            AddData3Radio("ConRat", rbConRat1, rbConRat2, rbConRat3)
            AddData3Radio("ConOthRodent", rbConOthRodent1, rbConOthRodent2, rbConOthRodent3)
            txtConOthRodentSpe.Text = .Rows(0)("ConOthRodentSpe") & ""
            AddData3Radio("ConStrayAni", rbConStrayAni1, rbConStrayAni2, rbConStrayAni3)
            txtConStrayAniSpe.Text = .Rows(0)("ConStrayAniSpe") & ""
            AddData3Radio("InsectBite", rbInsectBite1, rbInsectBite2, rbInsectBite3)
            AddData3Radio("ConMosquitoes", rbConMosquitoes1, rbConMosquitoes2, rbConMosquitoes3)
            AddData3Radio("ConFebrileFamily", rbConFebrileFamily1, rbConFebrileFamily2, rbConFebrileFamily3)
            AddData3Radio("ConFebrileCowork", rbConFebrileCowork1, rbConFebrileCowork2, rbConFebrileCowork3)
            AddData3Radio("ConFebrileNeighbor", rbConFebrileNeighbor1, rbConFebrileNeighbor2, rbConFebrileNeighbor3)
            AddData3Radio("WentAbroad", rbWentAboard1, rbWentAboard2, rbWentAboard3)
            If IsDBNull(.Rows(0)("WentAbroad")) OrElse .Rows(0)("WentAbroad") <> 1 Then
                SetComboDate(cboAbroadWhenDay, cboAbroadWhenMonth, cboAboradWhenYear)
            Else
                If String.IsNullOrEmpty(.Rows(0)("AbroadWhen").ToString) Then
                    SetComboDate(cboAbroadWhenDay, cboAbroadWhenMonth, cboAboradWhenYear)
                Else
                    cboAbroadWhenDay.Text = .Rows(0)("AbroadWhen").day
                    cboAbroadWhenMonth.Text = MonthName(.Rows(0)("AbroadWhen").month)
                    cboAboradWhenYear.Text = .Rows(0)("AbroadWhen").year + 543
                End If
            End If
            txtAboradWhere.Text = .Rows(0)("AbroadWhere") & ""
            AddData3Radio("VaccineJE", rbVaccineJE1, rbVaccineJE2, rbVaccineJE3)

            If IsDBNull(.Rows(0)("VaccineJE")) OrElse .Rows(0)("VaccineJE") <> 1 Then
                SetComboDate(cboLastJEDay, cboLastJEMonth, cboLastJEYear)
            Else
                If String.IsNullOrEmpty(.Rows(0)("LastJEDate").ToString) Then
                    SetComboDate(cboLastJEDay, cboLastJEMonth, cboLastJEYear)
                Else
                    cboLastJEDay.Text = .Rows(0)("LastJEDate").day
                    cboLastJEMonth.Text = MonthName(.Rows(0)("LastJEDate").month)
                    cboLastJEYear.Text = .Rows(0)("LastJEDate").year + 543
                End If
            End If
            
            AddData4Radio("LastJE", rbLastJE1, rbLastJE2, rbLastJE3, rbLastJE4)
            AddData2Radio("JEInfoSource", rbJEInfoSource1, rbJEInfoSource2)
            cboOccupation1.Text = IIf(IsDBNull(.Rows(0)("Occupation1")), 0, .Rows(0)("Occupation1"))
            txtOccupationOther1.Text = .Rows(0)("OccupationOther1") & ""
            cboOccupation2.Text = IIf(IsDBNull(.Rows(0)("Occupation2")), 0, .Rows(0)("Occupation2"))
            txtOccupationOther2.Text = .Rows(0)("OccupationOther2") & ""
            txtHematocrit.Text = .Rows(0)("Hematocrit") & ""
            chkHematocrit_ND.Checked = IIf(IsDBNull(.Rows(0)("Hematocrit_ND")), False, .Rows(0)("Hematocrit_ND"))
            txtPlatelet.Text = .Rows(0)("Platelet") & ""
            chkPlatelet_ND.Checked = IIf(IsDBNull(.Rows(0)("Platelet_ND")), False, .Rows(0)("Platelet_ND"))
            txtWBCCount.Text = .Rows(0)("WBCCount") & ""
            chkWBCCount_ND.Checked = IIf(IsDBNull(.Rows(0)("WBCCount_ND")), False, .Rows(0)("WBCCount_ND"))

            txtNEUTROPHIL.Text = .Rows(0)("NEUTROPHIL") & ""
            chkNEUTROPHIL_ND.Checked = IIf(IsDBNull(.Rows(0)("NEUTROPHIL_ND")), False, .Rows(0)("NEUTROPHIL_ND"))
            txtLYMPHOCYTE.Text = .Rows(0)("LYMPHOCYTE") & ""
            chkLYMPHOCYTE_ND.Checked = IIf(IsDBNull(.Rows(0)("LYMPHOCYTE_ND")), False, .Rows(0)("LYMPHOCYTE_ND"))

            txtMONOCYTE.Text = .Rows(0)("MONOCYTE") & ""
            chkMONOCYTE_ND.Checked = IIf(IsDBNull(.Rows(0)("MONOCYTE_ND")), False, .Rows(0)("MONOCYTE_ND"))

            txtEOSINOPHIL.Text = .Rows(0)("EOSINOPHIL") & ""
            chkEOSINOPHIL_ND.Checked = IIf(IsDBNull(.Rows(0)("EOSINOPHIL_ND")), False, .Rows(0)("EOSINOPHIL_ND"))

            txtWBCNameOther.Text = .Rows(0)("WBCNameOther") & ""
            txtWBCOtherResult.Text = .Rows(0)("WBCOtherResult") & ""
            txtBUN.Text = .Rows(0)("BUN") & ""
            chkBUN_ND.Checked = IIf(IsDBNull(.Rows(0)("BUN_ND")), False, .Rows(0)("BUN_ND"))

            txtALT.Text = .Rows(0)("ALT") & ""
            chkALT_ND.Checked = IIf(IsDBNull(.Rows(0)("ALT_ND")), False, .Rows(0)("ALT_ND"))

            txtAST.Text = .Rows(0)("AST") & ""
            chkAST_ND.Checked = IIf(IsDBNull(.Rows(0)("AST_ND")), False, .Rows(0)("AST_ND"))

            txtCreatinine.Text = .Rows(0)("Creatinine") & ""
            chkCreatinine_ND.Checked = IIf(IsDBNull(.Rows(0)("Creatinine_ND")), False, .Rows(0)("Creatinine_ND"))

            txtAlbumin.Text = .Rows(0)("Albumin") & ""
            chkAlbumin_ND.Checked = IIf(IsDBNull(.Rows(0)("Albumin_ND")), False, .Rows(0)("Albumin_ND"))

            AddData2Radio("UrineTaken", rbUrineTaken1, rbUrineTaken2)
            cboUrineTakenHow.SelectedValue = IIf(IsDBNull(.Rows(0)("UrineTAkenHow")), 0, .Rows(0)("UrineTAkenHow"))
            AddData2Radio("RBCMore5", rbRBCMore51, rbRBCMore52)
            txtWBCUrine.Text = .Rows(0)("WBCUrine") & ""
            chkWBCUrine_ND.Checked = IIf(IsDBNull(.Rows(0)("WBCUrine_ND")), False, .Rows(0)("WBCUrine_ND"))

            txtCSF_WBC.Text = .Rows(0)("CSF_WBC") & ""
            chkCSF_WBC_ND.Checked = IIf(IsDBNull(.Rows(0)("CSF_WBC_ND")), False, .Rows(0)("CSF_WBC_ND"))

            txtCSF_RBC.Text = .Rows(0)("CSF_RBC") & ""
            chkCSF_RBC_ND.Checked = IIf(IsDBNull(.Rows(0)("CSF_RBC_ND")), False, .Rows(0)("CSF_RBC_ND"))

            txtCSF_Protein.Text = .Rows(0)("CSF_Protein") & ""
            chkCSF_Protein_ND.Checked = IIf(IsDBNull(.Rows(0)("CSF_Protein_ND")), False, .Rows(0)("CSF_Protein_ND"))

            txtCSF_Glucose.Text = .Rows(0)("CSF_Glucose") & ""
            chkCSF_Glucose_ND.Checked = IIf(IsDBNull(.Rows(0)("CSF_Glucose_ND")), False, .Rows(0)("CSF_Glucose_ND"))

            cboCSF_GramStain.SelectedValue = IIf(IsDBNull(.Rows(0)("CSF_GramStain")), 0, .Rows(0)("CSF_GramStain"))
            chkCSF_GramStain_ND.Checked = IIf(IsDBNull(.Rows(0)("CSF_GramStain_ND")), False, .Rows(0)("CSF_GramStain_ND"))

            AddData2Radio("IsCultureBlood1", rbIsCultureBlood11, rbIsCultureBlood12)

            If IsDBNull(.Rows(0)("IsCultureBlood1")) OrElse .Rows(0)("IsCultureBlood1") <> 1 Then
                    SetComboDate(cboBloodCollect1Day, cboBloodCollect1Month, cboBloodCollect1Year)
                Else
                    If String.IsNullOrEmpty(.Rows(0)("BloodCollectDate1").ToString) Then
                    SetComboDate(cboBloodCollect1Day, cboBloodCollect1Month, cboBloodCollect1Year)
                Else
                    cboBloodCollect1Day.Text = .Rows(0)("BloodCollectDate1").day
                    cboBloodCollect1Month.Text = MonthName(.Rows(0)("BloodCollectDate1").month)
                    cboBloodCollect1Year.Text = .Rows(0)("BloodCollectDate1").year + 543
                End If
            End If

            mskBloodCollect1Time.Text = .Rows(0)("BloodCollectTime1").ToString & ""
            AddData2Radio("IsCultureUrine", rbIsCultureUrine1, rbIsCultureUrine2)

            If IsDBNull(.Rows(0)("IsCultureUrine")) OrElse .Rows(0)("IsCultureUrine") <> 1 Then
                    SetComboDate(cboUrineCollectDay, cboUrineCollectMonth, cboUrineCollectYear)
                Else
                    If String.IsNullOrEmpty(.Rows(0)("UrineCollectDate").ToString) Then
                    SetComboDate(cboUrineCollectDay, cboUrineCollectMonth, cboUrineCollectYear)
                Else
                    cboUrineCollectDay.Text = .Rows(0)("UrineCollectDate").day
                    cboUrineCollectMonth.Text = MonthName(.Rows(0)("UrineCollectDate").month)
                    cboUrineCollectYear.Text = .Rows(0)("UrineCollectDate").year + 543
                End If
            End If
            mskUrineCollectTime.Text = .Rows(0)("UrineCollectTime").ToString
            AddData2Radio("IsCultureCSF", rbIsCultureCSF1, rbIsCultureCSF2)
            If IsDBNull(.Rows(0)("IsCultureCSF")) OrElse .Rows(0)("IsCultureCSF") <> 1 Then

                SetComboDate(cboCSFCollectDay, cboCSFCollectDay, cboCSFCollectYear)
            Else
                If String.IsNullOrEmpty(.Rows(0)("CSFCollectDate").ToString) Then
                    SetComboDate(cboCSFCollectDay, cboCSFCollectDay, cboCSFCollectYear)
                Else
                    cboCSFCollectDay.Text = .Rows(0)("CSFCollectDate").day
                    cboCSFCollectMonth.Text = MonthName(.Rows(0)("CSFCollectDate").month)
                    cboCSFCollectYear.Text = .Rows(0)("CSFCollectDate").year + 543
                End If
            End If
            mskCSFCollectTime.Text = .Rows(0)("CSFCollectTime").ToString
            AddData2Radio("IsCultureBlood2", rbIsCultureBlood21, rbIsCultureBlood22)
            If IsDBNull(.Rows(0)("IsCultureBlood2")) OrElse .Rows(0)("IsCultureBlood2") <> 1 Then

                SetComboDate(cboBloodCollect2Day, cboBloodCollect2Month, cboBloodCollect2Year)
            Else
                If String.IsNullOrEmpty(.Rows(0)("BloodCollectDate2").ToString) Then
                    SetComboDate(cboBloodCollect2Day, cboBloodCollect2Month, cboBloodCollect2Year)
                Else
                    cboBloodCollect2Day.Text = .Rows(0)("BloodCollectDate2").day
                    cboBloodCollect2Month.Text = MonthName(.Rows(0)("BloodCollectDate2").month)
                    cboBloodCollect2Year.Text = .Rows(0)("BloodCollectDate2").year + 543
                End If
            End If

            mskBloodCollect2Time.Text = .Rows(0)("BloodCollectTime2").ToString
            AddData2Radio("DengueIGM", rbDengueIGM1, rbDengueIGM2)
            txtDENGUEIGM_Ini.Text = .Rows(0)("DengueIGM_Ini") & ""
            AddData3Radio("DengueIGMResult",rbDengueIGMResultPos, rbDengueIGMResultNeg, rbDengueIGMResultInd)
            AddData2Radio("DengueNS", rbDengueNS1, rbDengueNS2)
            txtDengueNS_Ini.Text = .Rows(0)("DengueNS_Ini") & ""
            AddData3Radio("DengueNSResult", rbDengueNSResultPos, rbDengueNSResultNeg, rbDengueNSResultInd)
            AddData2Radio("PneuUri", rbPneuUri1, rbPneuUri2)
            txtPneuUri_Ini.Text = .Rows(0)("PneuUri_Ini") & ""
            AddData3Radio("PneuUriResult", rbPneuUriResultPos, rbPneuUriResultNeg, rbPneuUriResultInd)
            AddData2Radio("BPSerum", rbBPSerum1, rbBPSerum2)
            txtBPSerum_Ini.Text = .Rows(0)("BPSerum_Ini") & ""
            AddData3Radio("BPSerumResult", rbBPSerumResultPos, rbBPSerumResultNeg, rbBPSerumResultInd)
            AddData2Radio("BPUrine", rbBPUrine1, rbBPUrine2)
            txtBPUrine_Ini.Text = .Rows(0)("BPUrine_Ini") & ""
            AddData3Radio("BPUrineResult", rbBPUrineResultpos, rbBPUrineResultNeg, rbBPUrineResultInd)
            AddData2Radio("BPSputum", rbBPSputum1, rbBPSputum2)
            txtBPSputum_Ini.Text = .Rows(0)("BPSputum_Ini") & ""
            AddData3Radio("BPSputumResult", rbBPSputumResultPos, rbBPSputumResultNeg, rbBPSputumResultInd)
            txtOtherRapidTest1.Text = .Rows(0)("OtherRapidTest1") & ""
            If IsDBNull(.Rows(0)("DischargeDate")) Then
                SetComboDate(cboDischargeDay, cboDischargeMonth, cboDischargeYear)
            Else
                cboDischargeDay.Text = .Rows(0)("DischargeDate").day
                cboDischargeMonth.Text = MonthName(.Rows(0)("DischargeDate").month)
                cboDischargeYear.Text = .Rows(0)("DischargeDate").year + 543
            End If
            AddData3Radio("Dengue", rbDengue1, rbDengue2, rbDengue3)
            AddData3Radio("Influenza", rbInfluenza1, rbInfluenza2, rbInfluenza3)
            AddData3Radio("Leptospirosis", rbLeptospirosis1, rbLeptospirosis2, rbLeptospirosis3)
            AddData3Radio("ScrubTyphus", rbScrubTyphus1, rbScrubTyphus2, rbScrubTyphus3)
            AddData3Radio("UpRespTract", rbUpRespTract1, rbUpRespTract2, rbUpRespTract3)
            AddData3Radio("UrineTract", rbUrineTract1, rbUrineTract2, rbUrineTract3)
            AddData3Radio("FeverUnknown", rbFeverUnknown1, rbFeverUnknown2, rbFeverUnknown3)
            AddData3Radio("InfectWound", rbInfectWound1, rbInfectWound2, rbInfectWound3)
            AddData3Radio("Typhoid", rbTyphoid1, rbTyphoid2, rbTyphoid3)
            AddData3Radio("Septicemia", rbSepticemia1, rbSepticemia2, rbSepticemia3)
            AddData3Radio("Gastroenteritis", rbGastroenteritis1, rbGastroenteritis2, rbGastroenteritis3)
            AddData3Radio("Melioidosis", rbMelioidosis1, rbMelioidosis2, rbMelioidosis3)
            AddData3Radio("Cellulitis", rbCellulitis1, rbCellulitis2, rbCellulitis3)
            AddData3Radio("Pneumonia", rbPneumonia1, rbPneumonia2, rbPneumonia3)
            AddData3Radio("Malaria", rbMalaria1, rbMalaria2, rbMalaria3)
            AddData3Radio("Chikungunya", rbChikungunya1, rbChikungunya2, rbChikungunya3)
            AddData3Radio("JE", rbJE1, rbJE2, rbJE3)
            AddData2Radio("Diabetes", rbDiabetes1, rbDiabetes2)
            AddData2Radio("Hypertension", rbHypertension1, rbHypertension2)
            AddData2Radio("HeartDisease", rbHeartDisease1, rbHeartDisease2)
            AddData2Radio("Asthma", rbAsthma1, rbAsthma2)
            AddData2Radio("CurSmoking", rbCurSmoking1, rbCurSmoking2)
            AddData2Radio("COPD", rbCOPD1, rbCOPD2)
            AddData2Radio("Cancer", rbCancer1, rbCancer2)
            txtCancerType.Text = .Rows(0)("CancerType") & ""
            AddData2Radio("Immunodeficiency", rbImmunodeficiency1, rbImmunodeficiency2)
            txtImmunoSpe.Text = .Rows(0)("ImmunoSpe") & ""
            AddData2Radio("KnownCase", rbHIV1, rbHIV2)
            AddData2Radio("HisTuber", rbHisTuber1, rbHisTuber2)
            AddData2Radio("ActTuber", rbActTuber1, rbActTuber2)
            AddData2Radio("Liver", rbLiver1, rbLiver2)
            AddData2Radio("Thyroid", rbThyroid1, rbThyroid2)
            AddData2Radio("Thalassemia", rbThalassemia1, rbThalassemia2)
            AddData2Radio("Anemia", rbAnemia1, rbAnemia2)
            AddData2Radio("ChroRenal", rbChroRenal1, rbChroRenal2)
            txtOtherCo1.Text = .Rows(0)("OtherCo1") & ""
            txtOtherCo2.Text = .Rows(0)("OtherCo2") & ""
            txtOtherCo3.Text = .Rows(0)("OtherCo3") & ""
            txtPrincipal1.Text = .Rows(0)("Principal1") & ""
            txtPrincipalICD10_1.Text = .Rows(0)("PrincipalICD10_1") & ""
            txtPrincipal2.Text = .Rows(0)("Principal2") & ""
            txtPrincipalICD10_2.Text = .Rows(0)("PrincipalICD10_2") & ""
            txtPrincipal3.Text = .Rows(0)("Principal3") & ""
            txtPrincipalICD10_3.Text = .Rows(0)("PrincipalICD10_3") & ""
            txtSecondary1.Text = .Rows(0)("Secondary1") & ""
            txtSecondaryICD10_1.Text = .Rows(0)("SecondaryICD10_1") & ""
            txtSecondary2.Text = .Rows(0)("Secondary2") & ""
            txtSecondaryICD10_2.Text = .Rows(0)("SecondaryICD10_2") & ""
            txtSecondary3.Text = .Rows(0)("Secondary3") & ""
            txtSecondaryICD10_3.Text = .Rows(0)("SecondaryICD10_3") & ""
            txtComplication1.Text = .Rows(0)("Complication1") & ""
            txtComplicationICD10_1.Text = .Rows(0)("ComplicationICD10_1") & ""
            txtComplication2.Text = .Rows(0)("Complication2") & ""
            txtComplicationICD10_2.Text = .Rows(0)("ComplicationICD10_2") & ""
            txtComplication3.Text = .Rows(0)("Complication3") & ""
            txtComplicationICD10_3.Text = .Rows(0)("ComplicationICD10_3") & ""
            AddData5Radio("DischargeStatus", rbDischargeStatus1, rbDischargeStatus2, rbDischargeStatus3, rbDischargeStatus4, rbDischargeStatus5)
            AddData6Radio("DischargeType", rbDischargeType1, rbDischargeType2, rbDischargeType3, rbDischargeType4, rbDischargeType5, rbDischargeType6)
            cboTransferHosp.Text = IIf(IsDBNull(.Rows(0)("TransferHospTo")), 0, .Rows(0)("TransferHospTo"))
            txtTransferHospOth.Text = .Rows(0)("TransferHospToOther") & ""
            txtDischargeTypeOther.Text = .Rows(0)("DischargeTypeOther") & ""
            AddData3Radio("Intubation", rbIntubation1, rbIntubation2, rbIntubation3)

        End With
    End Sub




    Private Sub rbIsCultureCSF1_CheckedChanged(sender As Object, e As EventArgs) Handles rbIsCultureCSF1.CheckedChanged
        culture(rbIsCultureCSF1, cboCSFCollectDay, cboCSFCollectMonth, cboCSFCollectYear, mskCSFCollectTime)


    End Sub

End Class