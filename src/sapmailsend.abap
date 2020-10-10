*&---------------------------------------------------------------------*
*& Report  ZSTUDENT_REPORT_MAIL_SEND
*&
*&---------------------------------------------------------------------*
*&
*&@Fatih Şengül
*&---------------------------------------------------------------------*

REPORT  ZSTUDENT_REPORT_MAIL_SEND

TABLES : ZP1770_T001.
TYPES : BEGIN OF ty_ogrenci,
  OGRNO TYPE ZP1770_T001-OGRNO,
  OGRNM TYPE ZP1770_T001-OGRNM,
  OGRLNM TYPE ZP1770_T001-OGRLNM,
  BLM TYPE ZP1770_T001-BLM,
  KTCN TYPE ZP1770_T001-KTCN,
  END OF ty_ogrenci.

DATA :  it_ogrenciler TYPE TABLE OF ty_ogrenci,
        wa_ogrenci TYPE ty_ogrenci.

DATA: it_ogrencicp TYPE STANDARD TABLE OF ty_ogrenci,
      it_changes TYPE STANDARD TABLE OF ty_ogrenci.

DATA :  wa_fieldcat   TYPE  slis_fieldcat_alv,
        it_fieldcat   TYPE  slis_t_fieldcat_alv,
        it_layout TYPE slis_layout_alv,
        form_top_of_page  TYPE slis_formname VALUE 'FORM_TOP_OF_PAGE',
        form_callback TYPE slis_formname VALUE 'USER_COMMAND'.
data: rt_extab type slis_t_extab.

***************send mail with alv definition*****************************************
DATA: IT_OBJBIN   TYPE STANDARD TABLE OF SOLISTI1,
      IT_OBJTXT   TYPE STANDARD TABLE OF SOLISTI1,
      IT_OBJPACK  TYPE STANDARD TABLE OF SOPCKLSTI1,
      IT_RECLIST  TYPE STANDARD TABLE OF SOMLRECI1,
      IT_OBJHEAD  TYPE STANDARD TABLE OF SOLISTI1.

DATA: WA_DOCDATA TYPE SODOCCHGI1,
      WA_OBJTXT  TYPE SOLISTI1,
      WA_OBJBIN  TYPE SOLISTI1,
      WA_OBJPACK TYPE SOPCKLSTI1,
      WA_OBJPACK2 TYPE SOPCKLSTI1,
      WA_RECLIST TYPE SOMLRECI1,
*     W_DOCUMENT_DATA  TYPE  SODOCCHGI1,
      OUT_CHAR(15),
      W_TAB_LINES TYPE I.
DATA : IT_SMTP TYPE BAPIADSMTP OCCURS 0 WITH HEADER LINE.
DATA : IT_BAPIRET TYPE BAPIRET2 OCCURS 0 WITH HEADER LINE.
DATA : IT_MSEG LIKE MSEG OCCURS 0 WITH HEADER LINE.

****************send mail with alv definition**************************************

SELECT-OPTIONS : s_ogrno FOR zp1770_t001-ogrno.

INITIALIZATION.

START-OF-SELECTION.
  perform set_pf_status USING RT_EXTAB.

  SELECT * FROM zp1770_t001 INTO CORRESPONDING FIELDS OF TABLE it_ogrenciler WHERE ogrno IN s_ogrno.

  PERFORM built_fcat.

END-of-SELECTION.

  PERFORM display_alv.


MODULE STATUS_0200 OUTPUT.
  SET PF-STATUS 'STATUS'.
ENDMODULE.

MODULE USER_COMMAND_0200 INPUT.
  CASE  sy-ucomm.
    WHEN 'ALV_BTN'.
      IF it_ogrenciler IS NOT INITIAL.
        PERFORM CREATE_MESSAGE.
        PERFORM SEND_MESSAGE.
      ENDIF.
    WHEN 'XLS_BTN'.
      perform send_mail_xls.
    when 'EXIT_BTN'.
      LEAVE TO SCREEN 0.
  ENDCASE.
ENDMODULE.



FORM set_pf_status USING RT_EXTAB.
  SET PF-STATUS 'GUI' EXCLUDING rt_extab.
ENDFORM.
form built_fcat.
  wa_fieldcat-fieldname = 'OGRNO'.
  wa_fieldcat-key = 'X'.
  wa_fieldcat-seltext_m = 'Öğrenci Numarası'.
  wa_fieldcat-hotspot = 'X'.
  APPEND wa_fieldcat TO it_fieldcat.
  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'OGRNM'.
  wa_fieldcat-seltext_m = 'OGRENCİ ADI'.
  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO it_fieldcat.
  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'OGRLNM'.
  wa_fieldcat-seltext_m = 'Öğrenci Soyadı'.
  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO it_fieldcat.
  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'BLM'.
  wa_fieldcat-seltext_m = 'BOLUM'.
  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO it_fieldcat.
  CLEAR wa_fieldcat.
  wa_fieldcat-fieldname = 'KTCN'.
  wa_fieldcat-seltext_m = 'Kitap Sayısı'.
  wa_fieldcat-edit = 'X'.
  APPEND wa_fieldcat TO it_fieldcat.
  CLEAR wa_fieldcat.

endform.


FORM display_alv.
  it_ogrencicp[] = it_ogrenciler[].
  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      i_callback_program       = sy-repid
      i_callback_user_command  = form_callback
      i_callback_top_of_page   = form_top_of_page
      i_callback_pf_status_set = 'SET_PF_STATUS'
      is_layout                = it_layout
      it_fieldcat              = it_fieldcat
    TABLES
      t_outtab                 = it_ogrenciler
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
            WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
ENDFORM.

FORM form_top_of_page.

  DATA: it_header TYPE slis_t_listheader,
        wa_header TYPE slis_listheader,
        lv_line LIKE wa_header-info,
        ld_lines TYPE i,
        ld_linesc(10) TYPE c.

  wa_header-typ  = 'H'.
  wa_header-info = 'Öğrenci Raporu'.
  APPEND wa_header TO it_header.
  CLEAR wa_header.

  wa_header-typ  = 'S'.
  wa_header-key = 'Date: '.
  CONCATENATE  sy-datum+6(2) '.'
               sy-datum+4(2) '.'
               sy-datum(4) INTO wa_header-info.
  APPEND wa_header TO it_header.
  CLEAR: wa_header.

  wa_header-typ  = 'S'.
  wa_header-key = 'User: '.
  CONCATENATE  sy-uname ' ' INTO wa_header-info.
  APPEND wa_header TO it_header.
  CLEAR: wa_header.

  DESCRIBE TABLE  it_ogrenciler LINES ld_lines.
  ld_linesc = ld_lines.
  CONCATENATE 'Total No. of Records Selected: ' ld_linesc
                    INTO lv_line SEPARATED BY space.
  wa_header-typ  = 'A'.
  wa_header-info = lv_line.
  APPEND wa_header TO it_header.
  CLEAR: wa_header, lv_line.

  CALL FUNCTION 'REUSE_ALV_COMMENTARY_WRITE'
    EXPORTING
      it_list_commentary = it_header.
  .
  CLEAR it_header.
ENDFORM.

FORM f_save_data.
  DATA : wa_ogrencicp    TYPE ty_ogrenci.
  DATA : wa_ogrenci_tmp TYPE ZP1770_T001.
  CLEAR it_changes[].
  LOOP AT it_ogrenciler INTO wa_ogrenci.
    READ TABLE it_ogrencicp INTO wa_ogrencicp INDEX sy-tabix.
    IF wa_ogrencicp NE wa_ogrenci.
      APPEND wa_ogrenci TO it_changes.
      MOVE-CORRESPONDING wa_ogrenci TO wa_ogrenci_tmp.
      MODIFY ZP1770_T001 FROM wa_ogrenci_tmp.
      IF sy-subrc EQ 0.
        MESSAGE 'Güncellendi' TYPE 'I'.
      ENDIF.
    ENDIF.
    CLEAR wa_ogrencicp.
  ENDLOOP.
ENDFORM.

FORM user_command  USING p_ucomm    LIKE sy-ucomm
                         p_selfield TYPE slis_selfield.
  CASE p_ucomm.
    WHEN '&IC1'.
      READ TABLE it_ogrenciler INTO wa_ogrenci INDEX p_selfield-tabindex.
      CASE P_SELFIELD-FIELDNAME.
        WHEN 'OGRN'.
          SET PARAMETER ID: 'OGR' FIELD P_SELFIELD-VALUE.
          CALL TRANSACTION 'MM03' AND SKIP FIRST SCREEN.
        WHEN 'MTART'.
          MESSAGE P_SELFIELD-VALUE TYPE 'S'.
      ENDCASE.
    WHEN '&UPDATE'.
      PERFORM f_save_data.
      MESSAGE 'Trying to save' TYPE 'S'.
    when '&SEND_MAIL'.
      perform send_mail_screen.
    WHEN '&BACK'.
      LEAVE PROGRAM.
  ENDCASE.
ENDFORM.

FORM send_mail_screen.
  call SCREEN 200 STARTING AT 55 5 ENDING AT 117 12.
endform.

**************************alv send***************************
form send_mail_alv.
  perform create_message.
  perform send_message.
endform.
******create message in html format*********
FORM CREATE_MESSAGE.


  CLEAR : IT_OBJTXT, OUT_CHAR.
  CLEAR: WA_OBJPACK.

  WRITE SY-DATUM DD/MM/YYYY TO OUT_CHAR.

  WA_DOCDATA-OBJ_NAME = 'Alert:For Bill Records'.
  CONCATENATE 'Date:' OUT_CHAR INTO WA_DOCDATA-OBJ_DESCR SEPARATED BY SPACE.

  WA_OBJTXT-LINE = '<html> <body>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<div style="border:1px solid;"> <div align="center">Öğrenci Raporu</div>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<table align="center">'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<tr>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<th style="background-color:#cfd5db;color:#000000;text-align:center;">Öğrenci Adı</th>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<th style="background-color:#cfd5db;color:#000000;text-align:center;">Öğrenci Soyadı</th>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<th style="background-color:#cfd5db;color:#000000;text-align:center;">Öğrenci Bölümü.</th>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<th style="background-color:#cfd5db;color:#000000;text-align:center;">Öğrenci Numarası</th>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<th style="background-color:#cfd5db;color:#000000;text-align:center;">Toplam Okunan Kitap Sayısı.</th>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

*     WA_OBJTXT-LINE = '<th style="background-color:#cfd5db;color:#000000;text-align:center;">Date</th>'.
*     APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '</tr>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  LOOP AT it_ogrenciler INTO wa_ogrenci.
    WA_OBJTXT-LINE = '<tr>'.
    APPEND WA_OBJTXT TO IT_OBJTXT.

    CONCATENATE '<td style="background-color:#c5eaee;border:1px solid;border-color:#72ccd6;text-align:right;">' wa_ogrenci-OGRNM '</td>' INTO WA_OBJTXT.
    APPEND WA_OBJTXT TO IT_OBJTXT.

    CONCATENATE '<td style="background-color:#f1f5fe;border:1px solid;border-color:#c4c4ff;text-align:right;">' wa_ogrenci-OGRLNM '</td>' INTO WA_OBJTXT.
    APPEND WA_OBJTXT TO IT_OBJTXT.

    CONCATENATE '<td style="background-color:#d5e3f2;border:1px solid;border-color:#6f9ed2;text-align:left;">' wa_ogrenci-BLM '</td>' INTO WA_OBJTXT.
    APPEND WA_OBJTXT TO IT_OBJTXT.

    CLEAR : OUT_CHAR.
    WRITE   wa_ogrenci-OGRNO TO OUT_CHAR.
    CONCATENATE '<td style="background-color:#d5e3f2;border:1px solid;border-color:#6f9ed2;text-align:right;">' OUT_CHAR '</td>' INTO WA_OBJTXT.
    APPEND WA_OBJTXT TO IT_OBJTXT.

    CLEAR : OUT_CHAR.
    WRITE : wa_ogrenci-KTCN TO OUT_CHAR.
    CONCATENATE '<td style="background-color:#fffdbf;border:1px solid;border-color:#fffb46;text-align:right">' OUT_CHAR '</td>' INTO WA_OBJTXT.
    APPEND WA_OBJTXT TO IT_OBJTXT.

*          CLEAR : OUT_CHAR.
*          WRITE : wa_ogrenci-MANDT TO OUT_CHAR.
*          CONCATENATE '<td style="background-color:#d5e3f2;border:1px solid;border-color:#6f9ed2;text-align:right">' OUT_CHAR '</td>' INTO WA_OBJTXT.
*          APPEND WA_OBJTXT TO IT_OBJTXT.

    WA_OBJTXT-LINE = '</tr>'.
    APPEND WA_OBJTXT TO IT_OBJTXT.

  ENDLOOP.

  WA_OBJTXT-LINE = '</table>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '</div>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '<div style="margin:30px 20px;font-size:12px;text-align:left;color:red;"> <b>Note:</b>- This is a SAP Training Test Mail, Please do not reply on this mail.</div></div>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.

  WA_OBJTXT-LINE = '</body> </html>'.
  APPEND WA_OBJTXT TO IT_OBJTXT.
  DESCRIBE TABLE IT_OBJTXT LINES W_TAB_LINES.
  READ TABLE IT_OBJTXT INTO WA_OBJTXT INDEX W_TAB_LINES.
  WA_DOCDATA-DOC_SIZE =
    ( W_TAB_LINES - 1 ) * 255 + STRLEN( WA_OBJTXT ).
  REFRESH IT_OBJPACK.
  CLEAR WA_OBJPACK-TRANSF_BIN.
  WA_OBJPACK-HEAD_START = 1.
  WA_OBJPACK-HEAD_NUM   = 0.
  WA_OBJPACK-BODY_START = 1.
  WA_OBJPACK-BODY_NUM   = W_TAB_LINES.
  WA_OBJPACK-DOC_TYPE   = 'HTML'.
  APPEND WA_OBJPACK TO IT_OBJPACK.

  CLEAR :IT_RECLIST,IT_SMTP,IT_BAPIRET.   " For BY passing
  REFRESH :IT_RECLIST,IT_SMTP,IT_BAPIRET.     " For BY passing UNAME

  CALL FUNCTION 'BAPI_USER_GET_DETAIL'
    EXPORTING
      USERNAME = SY-UNAME
    TABLES
      RETURN   = IT_BAPIRET
      ADDSMTP  = IT_SMTP.

  LOOP AT IT_SMTP .
    WA_RECLIST-RECEIVER = IT_SMTP-E_MAIL.
    WA_RECLIST-REC_TYPE = 'U'.
    APPEND WA_RECLIST TO IT_RECLIST.
    CLEAR  WA_RECLIST.
  ENDLOOP.

  WA_RECLIST-RECEIVER = 'ertugrul.cinar@detaysoft.com'.
  WA_RECLIST-REC_TYPE = 'U'.
*   WA_RECLIST-BLIND_COPY = 'X'.
  APPEND WA_RECLIST TO IT_RECLIST.
  CLEAR  WA_RECLIST.

  IF SY-UNAME EQ 'P1770'.
    CLEAR IT_RECLIST.       " For BY passing
    REFRESH IT_RECLIST.     " For BY passing
    CLEAR  WA_RECLIST.
    WA_RECLIST-RECEIVER = 'ertugrul.cinar@detaysoft.com'.
    WA_RECLIST-REC_TYPE = 'U'.
    APPEND WA_RECLIST TO IT_RECLIST.
    CLEAR  WA_RECLIST.
  ENDIF.
ENDFORM.
******create message in html format**************

************send message in html format**********
FORM SEND_MESSAGE.
  CALL FUNCTION 'SO_NEW_DOCUMENT_ATT_SEND_API1'
    EXPORTING
      DOCUMENT_DATA              = WA_DOCDATA
      PUT_IN_OUTBOX              = ' '
      COMMIT_WORK                = 'X'     "used from rel.6.10
    TABLES
      PACKING_LIST               = IT_OBJPACK
      OBJECT_HEADER              = IT_OBJHEAD
      CONTENTS_TXT               = IT_OBJTXT
      CONTENTS_BIN               = IT_OBJBIN
      RECEIVERS                  = IT_RECLIST
    EXCEPTIONS
      TOO_MANY_RECEIVERS         = 1
      DOCUMENT_NOT_SENT          = 2
      DOCUMENT_TYPE_NOT_EXIST    = 3
      OPERATION_NO_AUTHORIZATION = 4
      PARAMETER_ERROR            = 5
      X_ERROR                    = 6
      ENQUEUE_ERROR              = 7
      OTHERS                     = 8.

  IF SY-SUBRC NE 0.
    MESSAGE  'Mail sending failed' TYPE 'I'.
  ELSE.
    SUBMIT RSCONN01 WITH MODE = 'INT' WITH OUTPUT = ' ' AND RETURN.
    MESSAGE 'Mail sent successfully'  TYPE 'I'.
  ENDIF.

  WAIT UP TO 2 SECONDS.
  SUBMIT RSCONN01 WITH MODE = 'INT'
                WITH OUTPUT = ' '
                AND RETURN.
ENDFORM.
****************************send message in html format**************

****************************xls send*********************************
form send_mail_xls.

  DATA:
  gt_text           TYPE soli_tab,
  gt_attachment_hex TYPE solix_tab,
  gv_sent_to_all    TYPE os_boolean,
  gv_error_message  TYPE string,
  go_send_request   TYPE REF TO cl_bcs,
  go_recipient      TYPE REF TO if_recipient_bcs,
  go_sender         TYPE REF TO cl_sapuser_bcs,
  go_document       TYPE REF TO cl_document_bcs,
  gx_bcs_exception  TYPE REF TO cx_bcs.

  DATA:
  gc_subject     TYPE so_obj_des VALUE 'ABAP Email (attachment) with XLS',"içerik tanımı
  gc_email_to    TYPE adr6-smtp_addr VALUE 'ertugrul.cinar@detaysoft.com',"e-posta adresi
  gc_text        TYPE soli VALUE 'Öğrenci Raporu ektedir.',"mail texti
  gc_type_raw    TYPE so_obj_tp VALUE 'RAW',"döküman tipi
  gc_att_type    TYPE soodk-objtp VALUE 'XLS',"döküman tipi
  gc_att_subject TYPE sood-objdes VALUE 'Excel Dökümanı'."Nesne tanımı

*************** BINARY CONVERT****************************************
  DATA : LV_STRING TYPE STRING, "declare string
         LV_DATA_STRING TYPE STRING. "declare string
  LOOP AT IT_OGRENCILER INTO WA_OGRENCI.
    CONCATENATE WA_OGRENCI-OGRNM WA_OGRENCI-OGRLNM WA_OGRENCI-BLM INTO LV_STRING SEPARATED BY CL_ABAP_CHAR_UTILITIES=>HORIZONTAL_TAB.
    CONCATENATE LV_DATA_STRING LV_STRING INTO LV_DATA_STRING SEPARATED BY CL_ABAP_CHAR_UTILITIES=>NEWLINE.
    CLEAR: WA_OGRENCI, LV_STRING.
  ENDLOOP.

  DATA LV_XSTRING TYPE XSTRING .
**Xstring'e dönüştür.
  CALL FUNCTION 'HR_KR_STRING_TO_XSTRING'
    EXPORTING
      codepage_to      = '8300'
      UNICODE_STRING   = LV_DATA_STRING
*     OUT_LEN          =
    IMPORTING
      XSTRING_STREAM   = LV_XSTRING
    EXCEPTIONS
      INVALID_CODEPAGE = 1
      INVALID_STRING   = 2
      OTHERS           = 3.
  IF SY-SUBRC <> 0.
    IF SY-SUBRC = 1 .

    ELSEIF SY-SUBRC = 2 .
      WRITE:/ 'Geçersiz!!' .
    ENDIF.
  ENDIF.

  DATA: LIT_BINARY_CONTENT TYPE SOLIX_TAB.
***Xstring to binary
  CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
    EXPORTING
      BUFFER     = LV_XSTRING
    TABLES
      BINARY_TAB = LIT_BINARY_CONTENT.
*****************BINARY CONVERT*****************************************
  try.
      "Gönderme isteği oluştur
      go_send_request = cl_bcs=>create_persistent( ).
      "Gönderen Kişi
      go_sender = cl_sapuser_bcs=>create( sy-uname ).
      "İstek göndermek için gönderen eklenir
      go_send_request->set_sender( i_sender = go_sender ).
      "Gönderilen
      go_recipient = cl_cam_address_bcs=>create_internet_address( gc_email_to ).
      "Alıcı ekle
      go_send_request->add_recipient("mail giden alıcı
        EXPORTING
          i_recipient = go_recipient
          i_express   = abap_true
      ).

      "Email içerik
      APPEND gc_text TO gt_text."
      go_document = cl_document_bcs=>create_document(
                      i_type    = gc_type_raw
                      i_text    = gt_text
                      i_length  = '12'
                      i_subject = gc_subject ).


      "Maile Ek (Excel)
      clear gc_att_subject.
      CONCATENATE 'Öğrenci raporu ' sy-datum INTO gc_att_subject.
      go_document->add_attachment(
        EXPORTING
          i_attachment_type    = 'XLS'"gc_att_type"tipi
          i_attachment_subject = gc_att_subject"adı
          i_att_content_hex    = LIT_BINARY_CONTENT"içeriği!!!!
      ).

      "belgeyi ekle
      go_send_request->set_document( go_document )."dök. ekle

      "Mail'i gönder
      gv_sent_to_all = go_send_request->send( i_with_error_screen = abap_true ).
      IF gv_sent_to_all = abap_true.
        WRITE 'Bilgi: Email Gönderildi'.
      ENDIF.

      "Commitle
      COMMIT WORK.

      "Exception
    CATCH cx_bcs INTO gx_bcs_exception.
      gv_error_message = gx_bcs_exception->get_text( ).
      WRITE gv_error_message.
  ENDTRY.

  IF SY-SUBRC NE 0.
    MESSAGE  'Mail sending failed' TYPE 'I'.
  ELSE.
    SUBMIT RSCONN01 WITH MODE = 'INT' WITH OUTPUT = ' ' AND RETURN.
    MESSAGE 'Mail sent successfully'  TYPE 'I'.
  ENDIF.
endform.
****************************xls send*****************************