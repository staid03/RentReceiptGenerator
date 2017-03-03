#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#singleinstance , force

;Script for creating rent receipts

;Version	Date		Author		Notes
;	0.1		03-MAR-2017	Staid03		Initial

iniFile = .rentaldetails.ini		;file should be with the AHK script
receiptOutputFolder = D:

main:
{
	gosub , getVariablesFromIni
	gosub , goGui
}
return

goGui:
{
Gui , color , 00AAFF
Gui , Font , s11 CGreen
Gui , Add , Text , x50 y10 , Receipt Number
Gui , Add , Text , , Landlord Name
Gui , Add , Text , , Landlord Phone Number
Gui , Add , Text , , Landlord Email Address
Gui , Add , Text , , Rental Address
Gui , Add , Text , , Tenant Name
Gui , Add , Text , , Rent Amount
Gui , Add , Text , , Rent Amount Text
Gui , Font , CBlack
Gui , Add , Edit , w240 Right x220 y7 vReceiptNumber , %ReceiptNumberINI%
Gui , Add , Edit , w240 Right vLandLordName , %LandLordNameINI%
Gui , Add , Edit , w240 Right vLandLordPhoneNumber , %LandLordPhoneNumberINI%
Gui , Add , Edit , w240 Right vLandLordEmailAddress , %LandLordEmailAddressINI%
Gui , Add , Edit , w240 Right vRentalAddress , %RentalAddressINI%
Gui , Add , Edit , w240 Right vTenantName , %TenantNameINI%
Gui , Add , Edit , w240 Right vRentAmount , %RentAmountINI%
Gui , Add , Edit , w240 Right vRentAmountText , %RentAmountTextINI%
Gui , Add , Text , x10 , From:
Gui , Add , MonthCal , vStartRentReceiptDate , %ReceiptStartDate%
;Gui , Add , Text , x260 y231, To:
Gui , Add , Text , x260 y263, To:
Gui, Add , MonthCal , vEndRentReceiptDate , %ReceiptEndDate%
Gui , Add , Button , x10 w476 h40 , Go
Gui , show
}
Return 

ButtonGo:
{
	Gui , Submit , NoHide
	gosub , createText
	gosub , createReceipt
	gosub , updateINI
	gui , destroy
	sleep , 1000
	goto , main
}

Return

createText:
{
	finalLine = 				;reset finalLine
	formatTime , aDate , , dd-MMM-yyy
	formatTime , StartDate , %StartRentReceiptDate% , dd-MMM-yyy
	formatTime , EndDate , %EndRentReceiptDate% , dd-MMM-yyy
	
	borderline = ---------------------------------------------------------
	line = %borderline%
	gosub , addLine
	line = | %a_tab%%a_tab%%a_tab%Rental Receipt%a_tab%%a_tab%%a_tab%|	
	gosub , addLine
	;double-digit-format-checker
	ifless , ReceiptNumber , 10
	{
		ReceiptNumberSpaces = %ReceiptNumber%%a_space%%a_space%|
	}
	else
	{
		ReceiptNumberSpaces = %ReceiptNumber%%a_space%|
	}
	line = | Number:%a_tab%%a_tab%%a_tab%%a_tab%%a_tab%     %ReceiptNumberSpaces%
	gosub , addLine
	line = | Date Paid:%a_tab%%a_tab%%a_tab%%a_tab%   %aDate% %a_tab%|
	gosub , addLine
	line = | Address: %a_tab%%a_tab%%RentalAddress% %a_tab%|
	gosub , addLine
	line = | For: %a_tab%%a_tab%%a_tab%%a_tab%       %TenantName% %a_tab%|
	gosub , addLine
	line = | From: %a_tab%%a_tab%%a_tab%%a_tab%      %LandLordName% %a_tab%|
	gosub , addLine
	line = | Phone: %LandLordPhoneNumber%%a_tab%Email: %LandLordEmailAddress% %a_tab%|
	gosub , addLine
	line = | Amount: %a_tab%%a_tab%%a_tab%%a_tab%%a_tab%  $%RentAmount% %a_tab%|
	gosub , addLine
	line = | Amount Text: %a_tab%       %RentAmountText% Dollars %a_tab%|
	gosub , addLine
	line = | Rental Period From:%a_tab%%StartDate%    To: %EndDate% %a_tab%|
	; gosub , addLine
	; line = Date To: %EndDate%
	gosub , addLine
	line = |%a_tab%%a_tab%%a_tab%%a_tab%%a_tab%%a_tab%%a_tab%|
	gosub , addLine
	line = | Thank you%a_tab%%a_tab%%a_tab%%a_tab%%a_tab%%a_tab%|	
	gosub , addLine
	line = %borderline%
	gosub , addLine
}
Return

createReceipt:
{
	outfileReceiptName = %receiptOutputFolder%\Rental_Receipt_%EndRentReceiptDate%_Number_%ReceiptNumber%.txt
	ifexist , %outfileReceiptName%
	{
		formattime , xtime , , yyyyMMMdd_HHmmss
		filemove , %outfileReceiptName% , %outfileReceiptName%_%xtime%.txt
	}
	run , Notepad.exe
	sleep , 1000
	send , %finalLine%
	sleep , 200
	send , !F
	sleep , 200
	send , A
	sleep , 200
	send , %outfileReceiptName%
	sleep , 200
	send , {enter}
}
Return

addLine:
{
	finalLine = %finalLine%%line%`n
}
Return

getVariablesFromIni:
{
	iniRead , LandLordNameINI , %iniFile% , landlord_details , Name
	iniRead , LandLordPhoneNumberINI , %iniFile% , landlord_details , PhoneNumber
	iniRead , LandLordEmailAddressINI , %iniFile% , landlord_details , EmailAddress
	iniRead , TenantNameINI , %iniFile% , tenant_details , Name
	;iniRead , TenantPhoneNumber , %iniFile% , tenant_details , PhoneNumber
	;iniRead , TenantEmailAddress , %iniFile% , tenant_details , EmailAddress
	iniRead , RentalAddressINI , %iniFile% , house_details , Address
	iniRead , RentAmountINI , %iniFile% , Other , TypicalRentAmount
	iniRead , RentAmountTextINI , %iniFile% , Other , TypicalRentAmountText
	iniRead , ReceiptNumberINI , %iniFile% , Other , ReceiptNumber
	iniRead , LastRentalDateINI , %iniFile% , last_rental_date , DateEntered
	ReceiptStartDate = %LastRentalDateINI%	
	ReceiptEndDate = %LastRentalDateINI%
	EnvAdd , ReceiptStartDate , 1 , D
	EnvAdd , ReceiptEndDate , 14 , D
}
Return 

updateINI:
{
	ReceiptNumber++	
	iniWrite , %ReceiptNumber%, %iniFile% , Other , ReceiptNumber
	iniWrite , %EndRentReceiptDate%, %iniFile% , last_rental_date , DateEntered
}
Return 