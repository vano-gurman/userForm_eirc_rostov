'�������� �. �.
'���������������� ����� ��� �������������� ��������

'SYSTEM
'#include GLOBAL_VBS\SYSTEM\LibIncludeSv.vbs
'#include GLOBAL_VBS\SYSTEM\Class\cSysExBusiness.vbs

'PS
'#include GLOBAL_VBS\PS\Class\cPsExData.vbs

Sub ExGetParamFormEIRC_Rostov(objParam, objParamOut)

   Dim objParamGlobal
   Dim objComExVendor
   Dim objParamVendor

   if Not(Scripter.ExistParameter("objParamGlobal")) then
      Set objParamGlobal = CreateObject("Lib2.IUbsParam")
      Scripter.Parameter("objParamGlobal") = objParamGlobal
   else
      Set objParamGlobal = Scripter.Parameter("objParamGlobal")
   end if

   
    '������� ������� "���������"
    objParamOut.Parameter("cmdRun")  = "ExSetParamControlEIRC_2"

    Dim arrControl

		arrControl = InitElement()

    Dim arrData
    Dim i
    Dim arrValue
    ReDim arrValue(0)

    '������ �������� �������� ����������
    '��������������� �� ����������� ������ �������� ����������
    
    objParamOut.Parameter("Controls") = arrControl

    arrValue(0) = "���"
    objParamOut.Parameter("0") = arrValue
    arrValue(0) = "����"
    objParamOut.Parameter("2") = arrValue
    arrValue(0) = "�����:           �����"
    objParamOut.Parameter("4") = arrValue
    arrValue(0) = "�����"
    objParamOut.Parameter("6") = arrValue
    arrValue(0) = "���"
    objParamOut.Parameter("8") = arrValue
    arrValue(0) = "��������"
    objParamOut.Parameter("10") = arrValue
    arrValue(0) = "��������"
    objParamOut.Parameter("12") = arrValue
    arrValue(0) = "����� �� ���������"
    objParamOut.Parameter("14") = arrValue
    arrValue(0) = "��� ��"
    objParamOut.Parameter("16") = arrValue
    arrValue(0) = "��� �������"
    objParamOut.Parameter("17") = arrValue
    arrValue(0) = "���. ����� ��"
    objParamOut.Parameter("18") = arrValue
    arrValue(0) = "���������"
    objParamOut.Parameter("19") = arrValue
    arrValue(0) = "��� ������"
    objParamOut.Parameter("20") = arrValue
    arrValue(0) = "����� �� ������"
    objParamOut.Parameter("21") = arrValue
    arrValue(0) = "����� ���������"
    objParamOut.Parameter("22") = arrValue
    arrValue(0) = "�����"
    objParamOut.Parameter("24") = arrValue
    

End Sub


Sub ExSetParamControlEIRC_1(objParam, objParamOut)

	Dim arrData
	Dim i
	Dim arrValue
	ReDim arrValue(0)

	Dim strSQL, RSC
	
	Dim arrControl
		arrControl = InitElement()

	ReDim arrValue(0)


	Dim objOdbc
	Set objOdbc    = UBSCreateLib("UBSPublic3.1", "UBSPublic3", Scripter)
	
	Dim acc_r
		acc_r = objParam.Parameter("�������� �������� 3")
	
	arrValue(0) = ""
	for i=1 to 15 step 2
		if (i <> 3) Then objParamOut.Parameter(CStr(i)) = arrValue End If
	Next
	for i=26 to 26+4*6+2*20
		objParamOut.Parameter(CStr(i)) = arrValue
	Next
	
	'strSQL = "SELECT id_object FROM MINB_LIC_ACC_ADDFL WHERE field_string = '" & acc_r & "' AND id_field = " &_
	'	"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '������� ����'	)"
	strSQL = "	SELECT AD.id_object FROM MINB_LIC_ACC_ADDFL AD " &_
			"	JOIN MINB_LIC_ACC A ON AD.id_object = A.id " &_
			"	WHERE AD.field_string = '" & acc_r & "' AND A.ID_CONTRACT = " & Trim(objParam.Parameter("IDCONTRACT")) &_
			"	AND id_field =  " &_
			"	(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '������� ����'	)"
	GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
	
	Dim id_obj ' �� ��, ��� id_object
	if (TypeName(RSC) <> "Empty") Then
		id_obj = RSC(0,0)
	Else
		id_obj = 0
	End If
	
	Dim TmpStr
	
	Dim vParamI
	vParamI = objParam.Parameter("varParamIn")
	
	
	Dim q
	
	strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'ID ��������'	)"
	GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
	if (TypeName(RSC) = "Empty") Then
		' ������: ��� ������ �������!
		objParamOut.Parameter("���������") = "������!! ��� ������ �������!"
   		objParamOut.Parameter("���������") = False
   		exit sub
	ElseIf (Trim(RSC(0,0)) <> Trim(objParam.Parameter("IDCONTRACT"))) Then
		
		' ������: ������ ������� ��������
		
		objParamOut.Parameter("���������") = "������! ������ ����������� ������� ��������!"
   		objParamOut.Parameter("���������") = False
		exit sub
	Else
	
		
		
		' ���
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'NAME'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("1") = arrValue
		End If

		' �����
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'TOWN'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("5") = arrValue
		End If
		
		' �����
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'STREET'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("7") = arrValue
		End If
		
		' ���
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'HOUSE'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("9") = arrValue
		End If
		
		' ��������
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'BLOCK'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("11") = arrValue
		End If 
		
		' ��������
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = 'APARTMENT'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("13") = arrValue
		End If
		
		' ����� �� ���������
		dim SumPoKvi
			SumPoKvi = "0.0"
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '����� �� ���������'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			SumPoKvi = RSC(0,0)
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("15") = arrValue
		End If
		
		' ����� ���������
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '����� ���������'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			arrValue(0) = RSC(0,0)
			objParamOut.Parameter("23") = arrValue
		End If
		
	'-----------------------------------------------------
	'27 � 39

		Dim PU_count, Service_count
		
		strSQL = "SELECT COUNT(*) FROM MINB_LIC_ACC_ADDFL_ARRAY WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '���������� ����� ��'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			PU_count = RSC(0,0)
		End If

		strSQL = "SELECT COUNT(*) FROM MINB_LIC_ACC_ADDFL_ARRAY WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '��� ������'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
			Service_count = RSC(0,0)
		End If
		
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL_ARRAY WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '����� ��������'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
		For i=0 to PU_Count-1
			arrValue(0) = RSC(0,i)
			objParamOut.Parameter(CStr(26+i*4)) = arrValue
		Next
		End If
		
		strSQL = _ 
			"SELECT name_field FROM " &_
			"MINB_LIC_ACC_ADDFL_ARRAY a JOIN MINB_LIC_ACC_ADDFL_ARRAY_DIC_1 d " &_
			"ON field_string = d.id_field " &_
			"WHERE id_object = " & id_obj & " AND a.id_field = 14" '&_
	'	"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '��� �������'	)"

		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
		For i=0 to PU_Count-1
			arrValue(0) = RSC(0,i)
			objParamOut.Parameter(CStr(27+i*4)) = arrValue
		Next
		End If
		
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL_ARRAY WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '���������� ����� ��'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
		For i=0 to PU_Count-1
			arrValue(0) = RSC(0,i)
			objParamOut.Parameter(CStr(28+i*4)) = arrValue
		Next
		End If
		
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL_ARRAY WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '�������� ���������'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
		For i=0 to PU_Count-1
			arrValue(0) = RSC(0,i)
			objParamOut.Parameter(CStr(29+i*4)) = arrValue
		Next	
		End If				
		'Next
		
		strSQL = _ 
			"SELECT name_field FROM " &_
			"MINB_LIC_ACC_ADDFL_ARRAY a JOIN MINB_LIC_ACC_ADDFL_ARRAY_DIC_2 d " &_
			"ON field_string = d.id_field " &_
			"WHERE id_object = " & id_obj & " AND a.id_field = 17" '&_
		'"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '��� ������'	)"
		
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
		For i=0 to Service_Count-1
			arrValue(0) = RSC(0,i)
			objParamOut.Parameter(CStr(50 + i*2)) = arrValue  
		Next
		End If
		
		Dim sum
		
		strSQL = "SELECT field_string FROM MINB_LIC_ACC_ADDFL_ARRAY WHERE id_object = " & id_obj & " AND id_field = " &_
		"(	SELECT id_field FROM MINB_LIC_ACC_ADDFL_DIC WHERE Name_field = '����� �� ������'	)"
		GlobalDataAccess.Read OBJODBC.DSN(GlobalUser.SourceName), strSQL, RSC
		if (TypeName(RSC) <> "Empty") Then
		For i=0 to Service_Count-1
			if RSC(0,i) <> "" Then
				sum = sum + CCur ( join ( split(RSC(0,i), "."), ",") )
				arrValue(0) = RSC(0,i)
			Else
				arrValue(0) = "0.00"
			End If
			objParamOut.Parameter(CStr(51 + i*2)) = arrValue 'CStr(25 + Pu_count*4 + i*2)) = arrValue
		Next
			
			arrValue(0) = sum
			objParamOut.Parameter(CStr(25)) = arrValue
		Else
			' ������ �����, ����� �������
			'err.Raise -1, objParamOut.Parameter("23")
			if (objParamOut.Parameter("23") = "") Then ' ����� ��������� �����
				'err.Raise -1, "������ �����, ����� �������"
				sum = CCur ( join ( split(SumPoKvi, "."), ",") )
				arrValue(0) = sum
				objParamOut.Parameter(CStr(25)) = arrValue
			End If
		End If
		
	End If
	
	objParamOut.Parameter("Controls") = arrControl
	objParamOut.Parameter("���������") = True

End Sub




Sub ExSetParamControlEIRC_2(objParam, objParamOut)

	Dim varParamOut
	varParamOut = objParam.Parameter("varParamIn")
	
	Dim i, k
	Dim TmpStr
	Dim ArrP(5)
	dim tempstr, temparr, l
	
	Dim TmpArr1(3,0)
		TmpArr1(0,0) = 0
		TmpArr1(1,0) = "A"
		TmpArr1(2,0) = 0
		TmpArr1(3,0) = 0
	for i=0 to 5
		ArrP(i) = TmpArr1
	Next
	
	Dim ArrS(19)
	Dim TmpArr2(1,0)
		TmpArr2(0,0) = "T"
		TmpArr2(1,0) = 0
	
	for i=0 to 19
		ArrS(i) = TmpArr2
	Next

	for i=0 to 5

			ArrP(i)(0,0) = objParam.Parameter("�������� �������� " & CStr(26+i*4))'1
			ArrP(i)(1,0) = objParam.Parameter("�������� �������� " & CStr(27+i*4))'"���"
			ArrP(i)(2,0) = objParam.Parameter("�������� �������� " & CStr(28+i*4))'2
			ArrP(i)(3,0) = objParam.Parameter("�������� �������� " & CStr(29+i*4))'56.00
			
	Next
	for i=0 to 19

			ArrS(i)(0,0) = objParam.Parameter("�������� �������� " & CStr(26+4*6+i*2))'"��"
			ArrS(i)(1,0) = (objParam.Parameter("�������� �������� " & CStr(27+4*6+i*2)))  '1
	
			If (not (isNumeric(ArrS(i)(1,0)))) Then 
				temparr = split(ArrS(i)(1,0),".")				
				for k = 0 to Ubound(temparr)
				   if k = Ubound(temparr) then
					  tempstr = tempstr & temparr(k)
				   else
					  tempstr = tempstr & temparr(k) & ","
				   end if					   
				next
				l = len(tempstr)
				tempstr = Mid(tempstr,1,l)
				ArrS(i)(1,0) = tempstr	
				tempstr = ""		
			End If

	Next
	
	dim PU1StartPlace
	dim Us1StartPlace
	for i = 0 to UBound(varParamOut,2)
		select case varParamOut(0,i)
			case "_������ ����� 1"
				PU1StartPlace = i
			case "_������ 1"
				Us1StartPlace = i
		end select
	next
	
'		0: <txtCodePayment> 
'		1: <txtCode> 
'		2: <txtComment> 
'		3: <txtBic> 
'		4: <AccKorr> 
'		5: <txtNameBank> 
'		6: <AccClient> 
'		7: <txtINN> 
'		8: <txtRecip> 
'		9: <cmbPurpose> 
'		10: <txtFIOPay> 
'		11: <txtINNPay> 
'		12: <txtAdressPay> 
'		13: <txtInfoClient> 
'		14: <AccClientPay> 
'		15: <txtNomerCardPay> 
'		16: <txtKSPayment> 
'		17: <txtKSRate> 
'		18: <txtKSNDS> 
'		19: <txtDateBegin> 
'		20: <txtDateEnd> 
'		21: <AccPay> 
'		22: <txtCheckSum> 
'		23: <cboCityCode> 
'		24: <cmbTariff> 
'		25: <cmbPhone> 
'		26: <������ �����������> 
'		27: <��� ��������� �������������> 
'		28: <����������� ���> 
'		29: <��� �����> 
'		30: <��������� ���������� �������> 
'		31: <��������� ������> 
'		32: <����� ���������� ���������> 
'		33: <���� ���������� ���������> 
'		34: <��� ���������� �������> 
'		35: <����> 
'		36: <curSummaRateSend> 
'		37: <curSummaTotal> 
'		38: <curPeny> 
'		39: <curSumma> 
'		40: <����> 
'		41: <txtDayBeg> 
'		42: <txtMonthBeg> 
'		43: <txtYearBeg> 
'		44: <txtDayEnd> 
'		45: <txtMonthEnd> 
'		46: <txtYearEnd> 
'		47: <IdClient> 
'		48: <IdContract> 

'		49: <_������ ����� 1> 
'		50: <_������ ����� 2> 
'		51: <_������ ����� 3> 
'		52: <_������ ����� 4> 
'		53: <_������ ����� 5> 

'		54: <_������ 1> 
'		55: <_������ 2> 
'		56: <_������ 3> 
'		57: <_������ 4> 
'		58: <_������ 5> 
'		59: <_������ 6> 
'		60: <_������ 7> 
'		61: <_������ 8> 
'		62: <_������ 9> 
'		63: <_������ 10> 
'		64: <_������ 11> 
'		65: <_������ 12> 
'		66: <_������ 13> 
'		67: <_������ 14> 
'		68: <_������ 15> 
'		69: <_������ 16> 
'		70: <_������ 17> 
'		71: <_������ 18> 
'		72: <_������ 19> 
'		73: <_������ 20> 

'		74: <_������� ����> 

	for i = 0 to UBound(varParamOut,2)
	
	select case varParamOut(0,i)
		case "txtFIOPay"
			varParamOut(1,i) = objParam.Parameter("�������� �������� 1")
		case "_������� ����"
			varParamOut(1,i) = objParam.Parameter("�������� �������� 3")
		case "txtAdressPay"
		
			varParamOut(1,i) = _
				objParam.Parameter("�������� �������� 5") &_
				"," & objParam.Parameter("�������� �������� 7") &_
				"," & objParam.Parameter("�������� �������� 9")
			if (objParam.Parameter("�������� �������� 11") <> "") Then
				varParamOut(1,i) = varParamOut(1,i) &_
					",�" & objParam.Parameter("�������� �������� 11")
			End If
			varParamOut(1,i) = varParamOut(1,i) & "," & objParam.Parameter("�������� �������� 13")
		
		case "curSumma"
			varParamOut(1,i) = objParam.Parameter("�������� �������� 25")
		
		case "_������ ����� 1", _
		"_������ ����� 2", _
		"_������ ����� 3", _
		"_������ ����� 4", _
		"_������ ����� 5", _
		"_������ ����� 6"

			if (ArrP(i-PU1StartPlace)(0,0) <> "") Then 
			
				varParamOut(1,i) = ArrP(i-PU1StartPlace) 
			
			For k=0 to Ubound(varParamOut(1,i),1)
				TmpStr = TmpStr & TypeName(varParamOut(1,i)(k,0)) & ", "
			Next
			TmpStr = TmpStr & "; '" & ArrP(i-PU1StartPlace)(0,0) & "'" & VbCrLf
			
			End If
			
		case "_������ 1", _
			"_������ 2", _
			"_������ 3", _
			"_������ 4", _	
			"_������ 5", _
			"_������ 6", _
			"_������ 7", _
			"_������ 8", _
			"_������ 9", _
			"_������ 10", _
			"_������ 11", _
			"_������ 12", _
			"_������ 13", _
			"_������ 14", _
			"_������ 15", _
			"_������ 16", _
			"_������ 17", _
			"_������ 18", _
			"_������ 19", _
			"_������ 20"
			' 54 = 0, 55 = "1"
			if (ArrS(i-Us1StartPlace)(0,0) <> "" AND isNumeric(ArrS(i-Us1StartPlace)(1,0))) Then 
				varParamOut(1,i) = ArrS(i-Us1StartPlace) 
			
			For k=0 to Ubound(varParamOut(1,i),1)
				TmpStr = TmpStr & TypeName(varParamOut(1,i)(k,0)) & ", "
			Next
			TmpStr = TmpStr & ";" & VbCrLf
			
			End If
		
		end select
	
	next
	
	objParamOut.Parameter("���������") = True
	objParamOut.Parameter("varParamOut") = varParamOut

End Sub


Sub Sum_Count(objParam, objParamOut)

	Dim i
	Dim arrValue
	ReDim arrValue(0)
	
	Dim arrControl
		arrControl = InitElement()

	ReDim arrValue(0)



	Dim PU_count
		PU_count = 6
	Dim Service_count
		Service_count = 20

	Dim sum 
		sum = 0.00
		
	Dim TmpStr
	
		For i=0 to Service_Count
			TmpStr = join(split( objParam.Parameter("�������� �������� " & Cstr(27 + 4*Pu_Count + i*2)), "." ),",") 
			if isNumeric(TmpStr) Then 'TmpStr <> "" Then
				sum = sum + CCur (TmpStr)
			Else
				objParam.Parameter("�������� �������� " & Cstr(27 + 4*Pu_Count + i*2)) = ""
			End If

		Next
			
			' ��������� ����� ���������!
			TmpStr = join(split( objParam.Parameter("�������� �������� 23"), "." ),",") 
			if isNumeric(TmpStr) Then '<> "" Then
				sum = sum + CCur (TmpStr)
			Else					' ���� ����� ��������� ������
				If (sum = 0) Then 	' � ��� ���� sum �� ��� ��� = 0, 
									' ��, ������ �����, ��� ������� ������ ����,
									' ������� ���� ����� �� ���� "����� �� ���������"
					TmpStr = join(split( objParam.Parameter("�������� �������� 15"), "." ),",") 
					sum = CCur (TmpStr)
					'err.Raise -1, TmpStr & " " & sum
				End If
			End If
		
			arrValue(0) = sum
			objParamOut.Parameter(CStr(25)) = arrValue



	objParamOut.Parameter("Controls") = arrControl
	objParamOut.Parameter("���������") = True
	
	
End Sub


Function InitElement()

	Const Height = 315
	Const General_Top = 360
	'Const General_Left
	Const General_Width = 5800
	Const V_Offset = 390
	Const V_label_ofs = 60
	
	Const Arr_Begin = 26

	Dim PU_count
		PU_count = 6
	Dim Service_count
		Service_count = 20

	Dim ArrLength
		ArrLength = Arr_Begin-1 + PU_count*4 + Service_count*2
	
    Dim arrControl(7, 320)
    
    
    Dim k
    
    Dim i    
       For i=0 To 14 Step 2
		arrControl(0,i) = "Label"
		arrControl(1,i)=i
		arrControl(6,i)=""
		arrControl(7,i)=True
	Next
	For i=16 To 22
		arrControl(0,i) = "Label"
		arrControl(1,i)=i
		arrControl(6,i)=""
		arrControl(7,i)=True
	Next
		arrControl(0,24) = "Label"
		arrControl(1,24)=24
		arrControl(6,24)=""
		arrControl(7,24)=True

    For i=1 To 15 Step 2 
		arrControl(0,i) = "TextBox"
		arrControl(1,i)=i
		arrControl(5,i)=Height
		arrControl(7,i)=True
	Next
	For i=23 To ArrLength
		if i=24 Then
			arrControl(0,i) = "Label"
			arrControl(6,i)=""
		Else
			arrControl(0,i) = "TextBox"
			arrControl(5,i)=Height
		End If
		arrControl(1,i)=i
		arrControl(7,i)=True
	Next

    '0
    '  ��� �������� 
    '    Label
    '    TextBox
    '    ComboBox
    '    CommandButton
    '1
    '  ���������� Tag
    '    ���������� ����� �������� ���������� (�������� ������ �� 1)
    '2
    '  ������� Left �������� ����������
    '    �����
    '3
    '  ������� Top �������� ����������
    '    �����
    '4
    '  ������� Width �������� ����������
    '    �����
    '5
    '  ������� Height �������� ����������
    '    �����
    '6
    '  ������� �� ������� �������� ���������� (������� ������, ����� �� ��������) ��� Label �� ��������
    '    �����
    '7
    '  Enabled
    '    True, False


' ���
	k=0

    arrControl(2,k)=170
    arrControl(3,k)=General_Top + V_Offset*0 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255

	k=k+1
	
    arrControl(2,k)=3200
    arrControl(3,k)= General_Top + V_Offset*0
    arrControl(4,k)= General_Width
    arrControl(6,k)=""
    
' ����
	k=k+1

    arrControl(2,k)=170
    arrControl(3,k)=General_Top + V_Offset*1 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255

	k=k+1
   
    arrControl(2,k)=3200
    arrControl(3,k)= General_Top + V_Offset*1
    arrControl(4,k)=General_Width
    arrControl(6,k)="ExSetParamControlEIRC_1"
    
' �����: �����
	k=k+1

    arrControl(2,k)=170 + 3000
    arrControl(3,k)=General_Top + V_Offset*2 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255

	k=k+1

    arrControl(2,k)=6100
    arrControl(3,k)= General_Top + V_Offset*2
    arrControl(4,k)=General_Width/2
    arrControl(6,k)=""
    
' �����
	k=k+1

    arrControl(2,k)=170 + 4000
    arrControl(3,k)=General_Top + V_Offset*3 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255

	k=k+1

    arrControl(2,k)=6100
    arrControl(3,k)= General_Top + V_Offset*3
    arrControl(4,k)= General_Width/2
    arrControl(6,k)=""
    
' ���
 	k=k+1
   
    arrControl(2,k)=170 + 4000
    arrControl(3,k)=General_Top + V_Offset*4 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255

	k=k+1
   
    arrControl(2,k)=6100
    arrControl(3,k)= General_Top + V_Offset*4
    arrControl(4,k)=General_Width/2
    arrControl(6,k)=""
    
' ��������
	k=k+1

    arrControl(2,k)=170 + 4000
    arrControl(3,k)=General_Top + V_Offset*5 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255
    arrControl(7,k)=FALSE

	k=k+1
  
    arrControl(2,k)=6100
    arrControl(3,k)= General_Top + V_Offset*5
    arrControl(4,k)=General_Width/2
    arrControl(6,k)=""
    'arrControl(7,k)=FALSE
    
' ��������
	k=k+1

    arrControl(2,k)=170 + 4000
    arrControl(3,k)=General_Top + V_Offset*6 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255

	k=k+1
   
    arrControl(2,k)=6100
    arrControl(3,k)= General_Top + V_Offset*6
    arrControl(4,k)=General_Width/2
    arrControl(6,k)=""
    
' ����� �� ��������� 14-15
	k=k+1

    arrControl(2,k)=170
    arrControl(3,k)=General_Top + V_Offset*7 + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255
	
	k=k+1

    arrControl(2,k)=3200
    arrControl(3,k)= General_Top + V_Offset*7
    arrControl(4,k)=General_Width
    arrControl(6,k)=""



'------------------------------------------------
' ��� ��
	k=k+1

    arrControl(2,k)=1000 + V_label_ofs
    arrControl(3,k)=General_Top + V_Offset*8 + V_label_ofs
    arrControl(4,k)=1295
    arrControl(5,k)=255
    
' ��� �������
	k=k+1

    arrControl(2,k)=3000 + V_label_ofs
    arrControl(3,k)=General_Top + V_Offset*8 + V_label_ofs
    arrControl(4,k)=1295
    arrControl(5,k)=255
    
' ���������� ����� �� 18
	k=k+1

    arrControl(2,k)=5000 + V_label_ofs
    arrControl(3,k)=General_Top + V_Offset*8 + V_label_ofs
    arrControl(4,k)=1295
    arrControl(5,k)=255
    
' ������� ��������� 19
	k=k+1

    arrControl(2,k)=7000 + V_label_ofs
    arrControl(3,k)=General_Top + V_Offset*8 + V_label_ofs
    arrControl(4,k)=1295
    arrControl(5,k)=255

' ��� ������ 20
	k=k+1

    arrControl(2,k)=1000 + V_label_ofs
    arrControl(3,k)=General_Top + V_Offset*( 9 + PU_Count ) + V_label_ofs
    arrControl(4,k)=1295
    arrControl(5,k)=255
    
' ����� �� ������ 21
	k=k+1

    arrControl(2,k)=3000 + V_label_ofs
    arrControl(3,k)=General_Top + V_Offset*( 9 + PU_Count ) + V_label_ofs
    arrControl(4,k)=1295
    arrControl(5,k)=255
    
' ����� ��������� 22
   	k=k+1

    arrControl(2,k)=170
    arrControl(3,k)=General_Top + V_Offset*(10.3 + PU_Count + Service_Count) + V_label_ofs
    arrControl(4,k)=2895
    arrControl(5,k)=255
  	
	k=k+1

    arrControl(2,k)=3200
    arrControl(3,k)= General_Top + V_Offset*(10.3 + PU_Count + Service_Count)
    arrControl(4,k)=General_Width
    arrControl(6,k)="Sum_Count"
    
'-------------------------------------------------
' ����� �����������
	k=k+1
	
	arrControl(2,k)= 6000
	arrControl(3,k)= General_Top + V_Offset*( 9 + Service_count + PU_Count) + V_label_ofs
	arrControl(4,k)= 2895
    arrControl(5,k)=255
	
	k=k+1
	
	arrControl(2,k)= 7000
	arrControl(3,k)= General_Top + V_Offset*( 9 + Service_count + PU_Count)
	arrControl(4,k)= General_Width/3
	arrControl(6,k)=""
	arrControl(7,k)=FALSE
     
'--------------------------------------------------------

	For i = 26 to (26 + (PU_count-1)*4) step 4
		arrControl(2,i)= 1000
		arrControl(3,i)= General_Top + V_Offset*( 9 + (i-26)/4 )
		arrControl(4,i)= General_Width/3
	    arrControl(6,i)=""
		
		arrControl(2,i+1)= 3000
		arrControl(3,i+1)= General_Top + V_Offset*( 9 + (i-26)/4 )
		arrControl(4,i+1)= General_Width/3
	    arrControl(6,i+1)=""
	    		
		arrControl(2,i+2)= 5000
		arrControl(3,i+2)= General_Top + V_Offset*( 9 + (i-26)/4 )
		arrControl(4,i+2)= General_Width/3
	    arrControl(6,i+2)=""
	    		
		arrControl(2,i+3)= 7000
		arrControl(3,i+3)= General_Top + V_Offset*( 9 + (i-26)/4 )
		arrControl(4,i+3)= General_Width/3
	    arrControl(6,i+3)=""
	Next
	
	For i = 26 + 4*PU_count to (26 + 4*PU_count + (Service_count-1)*2) step 2
		arrControl(2,i)= 1000
		arrControl(3,i)= General_Top + V_Offset*( 9 + (i-26 - 4*PU_count)/2 + PU_Count + 1)
		arrControl(4,i)= General_Width/3
	    arrControl(6,i)=""
		
		arrControl(2,i+1)= 3000
		arrControl(3,i+1)= General_Top + V_Offset*( 9 + (i-26 - 4*PU_count)/2 + PU_Count + 1)
		arrControl(4,i+1)= General_Width/3
	    arrControl(6,i+1)="Sum_Count"
	Next

    
    InitElement = arrControl

End Function
