Attribute VB_Name = "invalidate_items"
'--------------------------------------------------
' �p�ԕi���グ�邽�߂̃A�b�v���[�h�t�@�C���𐶐����܂��B
'
' 2013.05.27
'--------------------------------------------------

'--------------------------------------------------
' �萔
'--------------------------------------------------
Private Const SHEET_NAME_SETUP As String = "�y�p�ԁzCSV����" '�ݒ�V�[�g��
Private Const START_DATA_ROW As Integer = 3 '�f�[�^�J�n�s(Jan�R�[�h�̃��X�g) 
Private Const MSG_1 As String = "Amazon�݌ɂ�0�ɂ��Ă��݌ɘA���ɂ�蕜������\��������܂��B" & vbCrLf & "Crossmall�̍݌ɘA����OFF�ɂ��邱�Ƃ��Y��Ȃ��悤�����ӂ��������I"
Private Const MSG_2 As String = "�Ƀt�@�C�����o�͂��܂����B"
Private Const MSG_3 As String = "�f�[�^�̓��͂�����܂���B"
Private Const MSG_4 As String = "�Y��JAN�̏��i�}�X�^���ȉ��̃X�e�[�^�X�ɒu�������܂��B" & vbCrLf & vbCrLf & "��\��" & vbCrLf & "�����_ = 0" & vbCrLf & "���؁i�����s�j�t���O = ON"



'--------------------------------------------------
' �t�@�C���o�̓t�H�[�}�b�g�ݒ�
'--------------------------------------------------
'#ECS��\��#
Private Const FILE_NAME_ECS_STATUS As String = "�status.csv"
Private Const HEADER_ECS_INACTIVATE = ""
Private Const LINE_REPLACE_ECS_INACTIVATE As String = "[JAN],2"

'#ECS�摜UP#
Private Const FILE_NAME_ECS_IMAGES As String = "�images.csv"
Private Const HEADER_ECS_IMAGES = ""
Private Const LINE_REPLACE_ECS_IMAGES As String = "[JAN]"

'#ECS��\���������_#
Private Const FILE_NAME_ECS_THRESHOLD As String = "�items_ecs.csv"
Private Const HEADER_ECS_THRESHOLD = "product_code,status,sellout_flg,delive_order_threshold"
Private Const LINE_REPLACE_ECS_THRESHOLD As String = "[JAN],2,1,0"

'#�y�V �q��#
Private Const FILE_NAME_RAKUTEN As String = "�item.csv"
Private Const HEADER_RAKUTEN = "�R���g���[���J����,���i�Ǘ��ԍ�(���iURL),�q�Ɏw��"
Private Const LINE_REPLACE_RAKUTEN As String = "u,[JAN],1"

'#�y�V �݌ɖ�#
Private Const FILE_NAME_RAKUTEN_STOCK As String = "�item.csv"
Private Const HEADER_RAKUTEN_STOCK = "�R���g���[���J����,���i�Ǘ��ԍ��i���iURL�j,�q�Ɏw��,�݌Ƀ^�C�v,�݌ɐ�,�݌ɐ��\��,�݌ɖ߂��t���O,�݌ɐ؂ꎞ�̒�����t"
Private Const LINE_REPLACE_RAKUTEN_STOCK As String = "u,[JAN],0,1,0,0,0,0"

'#�y�V �폜#
Private Const FILE_NAME_RAKUTEN_DELETE As String = "�item.csv"
Private Const HEADER_RAKUTEN_DELETE = "�R���g���[���J����,���i�Ǘ��ԍ��i���iURL�j"
Private Const LINE_REPLACE_RAKUTEN_DELETE As String = "d,[JAN]"

'#Yahoo �݌ɖ�#
Private Const FILE_NAME_YAHOO_STOCK As String = "�yahoo_stock.csv"
Private Const HEADER_YAHOO_STOCK = "code,sub-code,quantity,mode"
Private Const LINE_REPLACE_YAHOO_STOCK As String = "[JAN],,0,"

'#Yahoo �폜#
Private Const FILE_NAME_YAHOO_DELETE As String = "�yahoo_delete.csv"
Private Const HEADER_YAHOO_DELETE = "path,name,code,price"
Private Const LINE_REPLACE_YAHOO_DELETE As String = "a,a,[JAN],1"

'#Amazon �폜#
Private Const FILE_NAME_AMAZON_DELETE As String = "�amazon_delete.txt"
Private Const HEADER_AMAZON_DELETE = "TemplateType=Health[tab]Version=2012.1130[tab]���̍s��Amazon���g�p���܂��̂ŕύX��폜���Ȃ��ł��������B" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab]" & vbCrLf
Private Const HEADER_AMAZON_DELETE2 = "���i�Ǘ��ԍ�[tab]���i��[tab]���i�R�[�h(JAN�R�[�h��)[tab]���i�R�[�h�̃^�C�v[tab]�u�����h��[tab]" & _
    "���[�J�[��[tab]���i�^�C�v[tab]�p�b�P�[�W���i��[tab]���[�J�[�^��[tab]���i�����̉ӏ�����1[tab]���i�����̉ӏ�����2[tab]" & _
    "���i�����̉ӏ�����3[tab]���i�����̉ӏ�����4[tab]���i�����̉ӏ�����5[tab]���i������[tab]�����u���E�Y�m�[�h1[tab]" & _
    "�����L�[���[�h1[tab]�����L�[���[�h2[tab]�����L�[���[�h3[tab]�����L�[���[�h4[tab]�����L�[���[�h5[tab]���i���C���摜URL[tab]" & _
    "�݌ɐ�[tab]���[�h�^�C��(�o�ׂ܂łɂ������Ɠ���)[tab]���i�̔̔����i[tab]�ʉ݃R�[�h[tab]���i�̃R���f�B�V����[tab]" & _
    "���i�̃R���f�B�V��������[tab]�o�i�҃J�^���O�ԍ�[tab]���ޗ��E����1[tab]���ޗ��E����2[tab]���ޗ��E����3[tab]���ʐ���[tab]" & _
    "�g�p��̒���[tab]���i�̗��p(����)���@[tab]�x��[tab]�@�K��̖Ɛӏ���[tab]�A�_���g���i[tab]�X�^�C����[tab]" & _
    "�����u���E�Y�m�[�h2[tab]�e�q�֌W�̎w��[tab]�e���i��SKU(���i�Ǘ��ԍ�)[tab]�e�q�֌W�̃^�C�v[tab]�o���G�[�V�����e�[�}[tab]" & _
    "�t���[�o�[[tab]�T�C�Y[tab]�J���[[tab]�J���[�}�b�v[tab]����[tab]���i�̌`��[tab]����@�\1[tab]����@�\2[tab]����@�\3[tab]" & _
    "����p�r�L�[���[�h1[tab]����p�r�L�[���[�h2[tab]�Ώ�[tab]�J���[�T���v���摜URL[tab]���i�̃T�u�摜URL1[tab]���i�̃T�u�摜URL2[tab]" & _
    "���i�̃T�u�摜URL3[tab]���i�̃T�u�摜URL4[tab]���i�̃T�u�摜URL5[tab]���i�̃T�u�摜URL6[tab]���i�̃T�u�摜URL7[tab]" & _
    "���i�̃T�u�摜URL8[tab]�����ŏ��d��[tab]�����ő�d��[tab]�����d�ʂ̒P��[tab]���i�̌��J��[tab]���i�̏d�ʂ̒P��[tab]" & _
    "���i�̏d��[tab]���i�̒����̒P��[tab]���i�̒���[tab]���i�̕�[tab]���i�̍���[tab]�z���d�ʂ̒P��[tab]�z���d��[tab]" & _
    "�\�񏤕i�̔̔��J�n��[tab]���[�J�[��]�������i[tab]�g�p���Ȃ��x�������@[tab]�z�������w��SKU���X�g[tab]�Z�[�����i[tab]" & _
    "�t���t�B�������g�Z���^�[ID[tab]�Z�[���J�n��[tab]�Z�[���I����[tab]�ő咍����[tab]���i�̓��ח\���[tab]�ő哯���\��[tab]" & _
    "�M�t�g���b�Z�[�W[tab]�M�t�g�[tab]���[�J�[�������~[tab]���i�R�[�h�Ȃ��̗��R[tab]�v���`�i�L�[���[�h1[tab]�v���`�i�L�[���[�h2[tab]" & _
    "�v���`�i�L�[���[�h3[tab]�v���`�i�L�[���[�h4[tab]�v���`�i�L�[���[�h5[tab]�A�b�v�f�[�g�E�폜[tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & vbCrLf
Private Const HEADER_AMAZON_DELETE3 = "sku[tab]title[tab]standard-product-id[tab]product-id-type[tab]brand[tab]manufacturer[tab]" & _
    "product_type[tab]item-package-quantity[tab]mfr-part-number[tab]bullet-point1[tab]bullet-point2[tab]bullet-point3[tab]" & _
    "bullet-point4[tab]bullet-point5[tab]description[tab]recommended-browse-node1[tab]search-terms1[tab]search-terms2[tab]" & _
    "search-terms3[tab]search-terms4[tab]search-terms5[tab]main-image-url[tab]quantity[tab]leadtime-to-ship[tab]" & _
    "item-price[tab]currency[tab]condition-type[tab]condition-note[tab]merchant-catalog-number[tab]ingredients1[tab]" & _
    "ingredients2[tab]ingredients3[tab]special-ingredients[tab]indications[tab]directions[tab]warnings[tab]" & _
    "legal-disclaimer[tab]is-adult-product[tab]style-name[tab]recommended-browse-node2[tab]parentage[tab]parent-sku[tab]" & _
    "relationship-type[tab]variation-theme[tab]flavor[tab]size[tab]color[tab]color-map[tab]scent[tab]item-form[tab]" & _
    "special-features1[tab]special-features2[tab]special-features3[tab]specific-uses-keywords1[tab]specific-uses-keywords2[tab]" & _
    "target-audience[tab]swatch-image-url[tab]other-image-url1[tab]other-image-url2[tab]other-image-url3[tab]" & _
    "other-image-url4[tab]other-image-url5[tab]other-image-url6[tab]other-image-url7[tab]other-image-url8[tab]" & _
    "minimum-weight-recommendation[tab]maximum-weight-recommendation[tab]weight-recommendation-unit-of-measure[tab]" & _
    "launch-date[tab]item-weight-unit-of-measure[tab]item-weight[tab]item-length-unit-of-measure[tab]item-length[tab]" & _
    "item-width[tab]item-height[tab]shipping-weight-unit-of-measure[tab]shipping-weight[tab]release-date[tab]msrp[tab]" & _
    "optional-payment-type-exclusion[tab]delivery-schedule-group-id[tab]sale-price[tab]fulfillment-center-id[tab]" & _
    "sale-from-date[tab]sale-through-date[tab]max-order-quantity[tab]restock-date[tab]max-aggregate-ship-quantity[tab]" & _
    "is-gift-message-available[tab]is-giftwrap-available[tab]is-discontinued-by-manufacturer[tab]registered-parameter[tab]" & _
    "platinum-keywords1[tab]platinum-keywords2[tab]platinum-keywords3[tab]platinum-keywords4[tab]platinum-keywords5[tab]" & _
    "update-delete[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]"
Private Const LINE_REPLACE_AMAZON_DELETE As String = "[JAN][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]delete[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab]"


'#Amazon �݌�0#
Private Const FILE_NAME_AMAZON_STOCK As String = "�amazon_stock.txt"
Private Const HEADER_AMAZON_STOCK = "TemplateType=Health[tab]Version=1.7/1.2.11[tab]This row for Amazon.com use only.  " & _
    "Do not modify or delete.[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & vbCrLf
Private Const HEADER_AMAZON_STOCK2 = "sku[tab]title[tab]standard-product-id[tab]product-id-type[tab]brand[tab]manufacturer[tab]" & _
    "product_type[tab]mfr-part-number[tab]bullet-point1[tab]bullet-point2[tab]bullet-point3[tab]" & _
    "bullet-point4[tab]bullet-point5[tab]description[tab]recommended-browse-node1[tab]search-terms1[tab]" & _
    "search-terms2[tab]search-terms3[tab]search-terms4[tab]search-terms5[tab]main-image-url[tab]" & _
    "quantity[tab]leadtime-to-ship[tab]item-price[tab]currency[tab]merchant-catalog-number[tab]" & _
    "ingredients1[tab]ingredients2[tab]ingredients3[tab]indications[tab]directions[tab]warnings[tab]" & _
    "legal-disclaimer[tab]is-adult-product[tab]recommended-browse-node2[tab]parentage[tab]parent-sku[tab]" & _
    "relationship-type[tab]variation-theme[tab]flavor[tab]count[tab]size[tab]color[tab]scent[tab]" & _
    "other-image-url1[tab]other-image-url2[tab]other-image-url3[tab]other-image-url4[tab]" & _
    "other-image-url5[tab]other-image-url6[tab]other-image-url7[tab]other-image-url8[tab]launch-date[tab]" & _
    "item-weight-unit-of-measure[tab]item-weight[tab]item-length-unit-of-measure[tab]item-length[tab]" & _
    "item-width[tab]item-height[tab]shipping-weight-unit-of-measure[tab]shipping-weight[tab]" & _
    "release-date[tab]msrp[tab]sale-price[tab]fulfillment-center-id[tab]sale-from-date[tab]" & _
    "sale-through-date[tab]max-order-quantity[tab]restock-date[tab]max-aggregate-ship-quantity[tab]" & _
    "is-gift-message-available[tab]is-giftwrap-available[tab]is-discontinued-by-manufacturer[tab]" & _
    "registered-parameter[tab]platinum-keywords1[tab]platinum-keywords2[tab]platinum-keywords3[tab]" & _
    "platinum-keywords4[tab]platinum-keywords5[tab]update-delete"
Private Const LINE_REPLACE_AMAZON_STOCK As String = "[JAN][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab]0[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab][tab]" & _
    "[tab][tab][tab][tab]"

'--------------------------------------------------
' CSV�o�͊֐��̌Ăяo��
'--------------------------------------------------
'#ECS��\��#
Sub createEcsCSVInactivate()
    Call createCSV(FILE_NAME_ECS_STATUS, HEADER_ECS_INACTIVATE, LINE_REPLACE_ECS_INACTIVATE)
End Sub

'#ECS�摜UP#
Sub createEcsCSVImages()
    Call createCSV(FILE_NAME_ECS_IMAGES, HEADER_ECS_IMAGES, LINE_REPLACE_ECS_IMAGES)
End Sub

'#ECS��\���������_#
Sub createEcsCSVInactivateOrderThreshold()
    '���ӊ���
    If (MsgBox(MSG_4, vbInformation Or vbOKCancel) = vbCancel) Then
        Exit Sub
    End If
    
    Call createCSV(FILE_NAME_ECS_THRESHOLD, HEADER_ECS_THRESHOLD, LINE_REPLACE_ECS_THRESHOLD)
End Sub


'#�y�V#
Sub createRakutenCSV()
    Call createCSV(FILE_NAME_RAKUTEN, HEADER_RAKUTEN, LINE_REPLACE_RAKUTEN)
End Sub

'#�y�V �݌ɖ�#
Sub createRakutenCSVStock()
    Call createCSV(FILE_NAME_RAKUTEN_STOCK, HEADER_RAKUTEN_STOCK, LINE_REPLACE_RAKUTEN_STOCK)
End Sub

'#�y�V �폜#
Sub createRakutenCSVDelete()
    Call createCSV(FILE_NAME_RAKUTEN_DELETE, HEADER_RAKUTEN_DELETE, LINE_REPLACE_RAKUTEN_DELETE)
End Sub

'#Yahoo! �݌�#
Sub createYahooCSVStock()
    Call createCSV(FILE_NAME_YAHOO_STOCK, HEADER_YAHOO_STOCK, LINE_REPLACE_YAHOO_STOCK)
End Sub

'#Yahoo! �폜#
Sub createYahooCSVDelete()
    Call createCSV(FILE_NAME_YAHOO_DELETE, HEADER_YAHOO_DELETE, LINE_REPLACE_YAHOO_DELETE)
End Sub

'#Amazon �폜#
Sub createAmazonCSVDelete()
    Dim strHeader As String
    Dim strData As String
    
    '[tab]���^�u�R�[�h�ɒu��������
    strHeader = Replace(HEADER_AMAZON_DELETE, "[tab]", vbTab) & _
                Replace(HEADER_AMAZON_DELETE2, "[tab]", vbTab) & _
                Replace(HEADER_AMAZON_DELETE3, "[tab]", vbTab)
    strData = Replace(LINE_REPLACE_AMAZON_DELETE, "[tab]", vbTab)
    
    Call createCSV(FILE_NAME_AMAZON_DELETE, strHeader, strData)
End Sub

'#Amazon �݌�#
Sub createAmazonCSVStock()
    Dim strHeader As String
    Dim strData As String
    
    'Amazon�݌ɂɂ��Ă̒��ӊ���
    If (MsgBox(MSG_1, vbInformation Or vbOKCancel) = vbCancel) Then
        Exit Sub
    End If
    
    '[tab]���^�u�R�[�h�ɒu��������
    strHeader = Replace(HEADER_AMAZON_STOCK, "[tab]", vbTab) & _
                Replace(HEADER_AMAZON_STOCK2, "[tab]", vbTab)
    strData = Replace(LINE_REPLACE_AMAZON_STOCK, "[tab]", vbTab)
    
    Call createCSV(FILE_NAME_AMAZON_STOCK, strHeader, strData)
End Sub


' --------------------------------------------------
' CSV�o�� ���C��
'
' strFileName �c �o�͂���t�@�C����
' strHeader �c �o�͂���w�b�_�s�i�w�b�_�s���Ȃ��ꍇ�͋�Ƃ���j
' strLineReplace �c �f�[�^�s�̃t�H�[�}�b�g [JAN]�̕�����JAN�R�[�h�ɒu�������
' --------------------------------------------------
Private Sub createCSV(strFileName As String, strHeader As String, strLineReplace As String)
    Dim intEndRow As Integer
    Dim strFilePath As String
    Dim i As Integer
    Dim strLine As String '�o�͍s�����p
    Dim IntFlNo As Integer '�t�@�C���I�[�v���p
    
    '�o�̓t�@�C����
    strFilePath = ActiveWorkbook.Path & strFileName
    
    '�f�[�^�s�̗L���̊m�F
    intEndRow = Worksheets(SHEET_NAME_SETUP).Cells(Rows.Count, 2).End(xlUp).Row
    If (intEndRow < START_DATA_ROW) Then
        Call MsgBox(MSG_3, vbExclamation)
        Exit Sub
    End If
    
    '�t�@�C���I�[�v��
    IntFlNo = FreeFile
    Open strFilePath For Output As #IntFlNo
    
    '�w�b�_�s
    If (strHeader <> "") Then
        Print #IntFlNo, strHeader
    End If
    
    '�t�@�C���o��
    For i = START_DATA_ROW To intEndRow
        strLine = ""
        strLine = Worksheets(SHEET_NAME_SETUP).Cells(i, 2).Value
        
        If (strLine <> "") Then
            strLine = Replace(strLineReplace, "[JAN]", strLine)
            Print #IntFlNo, strLine
        End If
    Next
    
    '�t�@�C���N���[�Y
    Close #IntFlNo
    
    
    '�t�@�C���o�͌チ�b�Z�[�W
    Call MsgBox(strFilePath & vbCrLf & MSG_2, vbInformation)
End Sub

