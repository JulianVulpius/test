create or replace PACKAGE PREISDATENBANK_PKG AS

  TYPE R_AUFTRAEGE is RECORD (auftrag_id number,
                              column_position number);
  TYPE T_AUFTRAEGE IS TABLE OF R_AUFTRAEGE;


  TYPE R_KENNUNG IS RECORD (kennung varchar2(50),
                            row_position number);
  TYPE T_KENNUNG IS TABLE OF R_KENNUNG;

  PROCEDURE Import_Muster(p_blob_id number,p_typ_id number,p_date date);
  PROCEDURE Import_Muster1(p_blob_id number,p_typ_id number,p_date date);
  PROCEDURE Import_Muster2(p_blob_id number,p_typ_id number,p_date date);

  PROCEDURE import_auftrege(p_blob_id number, p_typ_id number,p_region_id number,p_out out varchar2);

  PROCEDURE auswertung_to_excel_2(p_von date, p_bis date, p_lvs varchar2,p_regionen varchar2,p_liferant varchar2,p_trimm number default null,
                                    p_vergabesumme_von number default null, p_vergabesumme_bis number default null,
                                    p_user_id number);

  procedure auswertung_to_excel_leser(p_von date, p_bis date, p_lvs varchar2,p_regionen varchar2,p_liferant varchar2,p_trimm number default null,
                                      p_vergabesumme_von number default null, p_vergabesumme_bis number default null,
                                      p_user_id number);

  PROCEDURE export_ausschreibung_to_excel(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number);
  PROCEDURE export_ausschreibung_to_excel_unified(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number);
  PROCEDURE export_ausschreibung_to_excel_gaeb90(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number);

  FUNCTION print_param_worksheet(p_ws  IN OUT  xlsx_writer.book_r, p_id IN NUMBER ,p_style in number) return  xlsx_writer.book_r;
  FUNCTION print_param_worksheet(p_ws IN OUT xlsx_writer.book_r, p_von IN DATE, p_bis IN DATE, p_regionen IN VARCHAR2, p_lvs_raw IN VARCHAR2, p_liferant_raw IN VARCHAR2, 
  p_trimm number default null, p_vergabesumme_von number default null, p_vergabesumme_bis number default null) return xlsx_writer.book_r;

  FUNCTION GetNamespace(p_blob_id in number) return number;

  PROCEDURE SendMailAuswertung(p_user_id number,p_anhang blob,p_filename varchar2);
  PROCEDURE SendMailAuswertungFehler(p_user_id number,p_error varchar2);

END PREISDATENBANK_PKG;