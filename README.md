# CopyFee
โปรแกรมจัดทำใบแจ้งชำระค่าธรรมเนียมการทำสำเนาข้อมูลข่าวสาร (CopyFee) พัฒนาโดยศูนย์บริการร่วม กระทรวงทรัพยากรธรรมชาติและสิ่งแวดล้อม เมื่อปี พ.ศ. 2565 เพื่อให้เจ้าหน้าที่ศูนย์ข้อมูลข่าวสาร สำนักงานปลัดกระทรวงทรัพยากรธรรมชาติและสิ่งแวดล้อม (สป.ทส.) ใช้ในการตรวจสอบและออกใบแจ้งชำระค่าธรรมเนียมการทำสำเนาข้อมูลข่าวสาร และการรับรองถูกต้องของข้อมูลข่าวสาร ในกรณีที่ต้องเรียกเก็บค่าธรรมเนียมจากผู้ยื่นคำขอข้อมูลข่าวสาร ตามประกาศ สป.ทส. เรื่อง การเรียกค่าธรรมเนียมการขอสำเนา หรือขอสำเนาที่มีคำรับรองถูกต้องของข้อมูลข่าวสารของราชการ ลงวันที่  27 พฤศจิกายน พ.ศ. 2549

## วัตถุประสงค์
- [x] ประชาชนได้รับใบแจ้งชำระค่าธรรมเนียมอย่างถูกต้องและรวดเร็ว
- [x] เจ้าหน้าที่ออกใบแจ้งชำระค่าธรรมเนียมสะดวกและรวดเร็วยิ่งขึ้น

## วิธีใช้
1. Double Click ที่ไฟล์ CopyFee2.1.docm
2. กดปุ่ม Start เพื่อเข้าสู่หน้าโปรแกรม
3. กรอกข้อมูล ทุกช่อง
4. กดปุ่ม Done เพื่อยืนยันข้อมูล
5. กดปุ่ม Print เพื่อออกใบแจ้งชำระค่าธรรมเนียม
6. กดปุ่ม Reset เพื่อล้างข้อมูล
7. กดปุ่ม Exit เพื่อปิดโปรแกรม
8. เปิดไฟล์ใบแจ้งชำระค่าธรรมเนียม (PDF) เพื่อตรวจสอบความถูกต้อง
9. สั่งพิมพ์ และลงนาม

## Vesion History
2.1
- Payment methods included.
- Better document format

## Programming (Algorithm & Coding)
- Create a form as table in Microsoft Word
- Use Legacy tools in Developer menu
- Use VBA script 
- Transfer the filled text from all cells in the form to the invoice, via the Done button:
```
Private Sub CommandButton3_Click()
    MsgBox "Ready to print", , "Done!"
End Sub
```
- Print the invoice to a PDF copy, via the Print button:
```
Private Sub CommandButton1_Click()
    Application.PrintOut FileName:="", Range:=wdPrintRangeOfPages, Item:= _
        wdPrintDocumentWithMarkup, Copies:=1, Pages:="2", PageType:= _
        wdPrintAllPages, Collate:=True, Background:=True, PrintToFile:=False, _
        PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
        PrintZoomPaperHeight:=0
End Sub
```
- Use the Reset button to reset the form, protected by password: *****
```
Private Sub CommandButton2_Click()
    ActiveDocument.Unprotect Password:="*****"
    ActiveDocument.ResetFormFields
    ActiveDocument.Protect Type:=wdAllowOnlyFormFields, _
                    Password:="*****"
End Sub
```
- Auto-fill data into the invoice from each cell in the form
```
{ REF Name \* CHARFORMAT } 
{ REF Date \* CHARFORMAT } 
{ REF Page } 
{ REF Place } 
{ REF Officer \* CHARFORMAT } 
{ REF Position \* CHARFORMAT } 
```
- Read the amount of payment in Thai text
```
(={ e4\*bahttext })
```
- Date of today
```
{ TIME \@ "ว ดดดด ปปปป" }
```
- Show the amount of pages and places of approval with comma separating thousands
```
{ REF Page \#,0 \*MERGEFORMAT}
{ REF Place \#,0 \*MERGEFORMAT}
```
- Close the program, with the Exit button
```
Private Sub CommandButton5_Click()
   Application.Quit
End Sub
```
## Programmer
Monte Kietpwpan, Ph.D.,
Service Link Center, Ministry of Natural Resources and Environment
