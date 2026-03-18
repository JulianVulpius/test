create or replace PACKAGE BODY           PREISDATENBANK_PKG AS

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE Import_Muster(p_blob_id number, p_typ_id number,p_date date) AS
    v_namespace number;
  BEGIN

    select GetNamespace(p_blob_id) into v_namespace from dual;

    if v_namespace = 1 then
      Import_Muster1(p_blob_id,p_typ_id,p_date);
    end if;

    if v_namespace = 2 then
      Import_Muster2(p_blob_id,p_typ_id,p_date);
    end if;

  END Import_Muster;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE Import_Muster1(p_blob_id number, p_typ_id number,p_date date) AS

    v_kennung   pd_import_typ.kennung%type;
    v_name      pd_import_typ.name%type;

  BEGIN

    select  kennung,name
    into    v_kennung,v_name
    from    pd_import_typ
    where   id = p_typ_id;

    for i in 
            (
            SELECT  v_kennung code,v_name name,null mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x
            where   x.id = p_blob_id
            union all    
            SELECT  v_kennung || '.' || xt.code,xt.name,v_kennung mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2,xtd.name2,v_kennung || '.' || xt.code,null description,null,null
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                                   PASSING xt.subitem
                                   COLUMNS 
                                           code2        VARCHAR2(2000)  PATH '@RNoPart',
                                           name2        VARCHAR2(4000)  PATH 'LblTx',
                                           subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2||'.'||xtd2.code3,xtd2.name3,v_kennung || '.' || xt.code||'.'||xtd.code2,xtd2.desc3,
                    case 
                        when xtd2.name4 like '%MLV%' or xtd2.name4 like '%MVL%' then v_kennung||substr(xtd2.name4,instr(xtd2.name4,'MLV')+instr(xtd2.name4,'MVL')+7,16)||' '||xtd2.name5
                        when xtd2.name5 like '%MLV%' or xtd2.name5 like '%MVL%' then v_kennung||substr(xtd2.name5,instr(xtd2.name5,'MLV')+instr(xtd2.name5,'MVL')+7,16)
                        else v_kennung||'_'||xt.code||xtd.code2||xtd2.code3 
                    end code,
                    xtd2.einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                               PASSING xt.subitem
                               COLUMNS 
                                       code2        VARCHAR2(2000)  PATH '@RNoPart',
                                       name2        VARCHAR2(4000)  PATH 'LblTx',
                                       subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/Item'
                               PASSING xtd.subsubitem
                               COLUMNS 
                                       code3        VARCHAR2(200)  PATH '@RNoPart',
                                       name3        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt',
                                       name4        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                       name5        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                       desc3        VARCHAR2(4000) PATH 'Description/CompleteText/DetailTxt',
                                       einheit      VARCHAR2(20)   PATH 'QU'
                            ) xtd2
            where x.id = p_blob_id
            )
    loop
        MERGE INTO PD_MUSTER_LVS e
        USING ( SELECT  i.code id,p_typ_id typ,i.mparent mojp FROM dual) h
        ON (e.code = h.id and e.MUSTER_TYP_ID = h.typ and (e.PARENT_ID = h.mojp or (e.PARENT_ID is null and h.mojp is null)))
        WHEN MATCHED THEN
            UPDATE 
            SET     e.name = i.name,
                    e.DESCRIPTION = i.description,
                    e.position_kennung = trim(i.kennung),
                    e.einheit = i.einheit,
                    e.stand_datum = p_date
        WHEN NOT MATCHED THEN
            INSERT (parent_id, code,name,description,muster_typ_id,position_kennung,STAND_DATUM,einheit)
            VALUES (i.mparent, i.code,i.name,i.description,p_typ_id,trim(i.kennung),p_date,i.einheit);
    end loop;

    delete from pd_import_x83 where id = p_blob_id;
  exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_MUSTER: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM || ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
        delete from pd_import_x83 where id = p_blob_id;
        raise_application_error(-20000,SQLERRM);
  END Import_Muster1;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE Import_Muster2(p_blob_id number, p_typ_id number,p_date date) AS

    v_kennung   pd_import_typ.kennung%type;
    v_name      pd_import_typ.name%type;

  BEGIN

    select  kennung,name
    into    v_kennung,v_name
    from    pd_import_typ
    where   id = p_typ_id;

    for i in 
            (
            SELECT  v_kennung code,v_name name,null mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x
            where   x.id = p_blob_id
            union all    
            SELECT  v_kennung || '.' || xt.code,xt.name,v_kennung mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2,xtd.name2,v_kennung || '.' || xt.code,null description,null,null
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                    code     VARCHAR2(2000)  PATH '@RNoPart',
                                    name     VARCHAR2(4000)  PATH 'n:LblTx',
                                    subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
                                PASSING xt.subitem
                                COLUMNS 
                                    code2        VARCHAR2(2000)  PATH '@RNoPart',
                                    name2        VARCHAR2(4000)  PATH 'LblTx',
                                    subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2||'.'||xtd2.code3,xtd2.name3,v_kennung || '.' || xt.code||'.'||xtd.code2,xtd2.desc3,
                    case 
                        when xtd2.name4 like '%MLV%' or xtd2.name4 like '%MVL%' then v_kennung||substr(xtd2.name4,instr(xtd2.name4,'MLV')+instr(xtd2.name4,'MVL')+7,16)||' '||xtd2.name5
                        when xtd2.name5 like '%MLV%' or xtd2.name5 like '%MVL%' then v_kennung||substr(xtd2.name5,instr(xtd2.name5,'MLV')+instr(xtd2.name5,'MVL')+7,16)
                        else v_kennung||'_'||xt.code||xtd.code2||xtd2.code3 
                    end code,
                    xtd2.einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                    code     VARCHAR2(2000)  PATH '@RNoPart',
                                    name     VARCHAR2(4000)  PATH 'n:LblTx',
                                    subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
                                PASSING xt.subitem
                                    COLUMNS 
                                        code2        VARCHAR2(2000)  PATH '@RNoPart',
                                        name2        VARCHAR2(4000)  PATH 'LblTx',
                                        subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/Item'
                                PASSING xtd.subsubitem
                                    COLUMNS 
                                        code3        VARCHAR2(200)  PATH '@RNoPart',
                                        name3        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt',
                                        name4        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                        name5        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                        desc3        VARCHAR2(4000) PATH 'Description/CompleteText/DetailTxt',
                                        einheit      VARCHAR2(20)   PATH 'QU'
                            ) xtd2
            where x.id = p_blob_id
            )
    loop
        MERGE INTO PD_MUSTER_LVS e
        USING (SELECT i.code id,p_typ_id typ,i.mparent mojp FROM dual) h
        ON (e.code = h.id and e.MUSTER_TYP_ID = h.typ and (e.PARENT_ID = h.mojp or (e.PARENT_ID is null and h.mojp is null)))
        WHEN MATCHED THEN
            UPDATE 
            SET     e.name = i.name,
                    e.DESCRIPTION = i.description,
                    e.position_kennung = trim(i.kennung),
                    e.einheit = i.einheit,
                    e.stand_datum = p_date
        WHEN NOT MATCHED THEN
            INSERT (parent_id, code,name,description,muster_typ_id,position_kennung,STAND_DATUM,einheit)
            VALUES (i.mparent, i.code,i.name,i.description,p_typ_id,trim(i.kennung),p_date,i.einheit);
    end loop;

    delete from pd_import_x83 where id = p_blob_id;
  exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_MUSTER: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM ||
      ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
      delete from pd_import_x83 where id = p_blob_id;
      raise_application_error(-20000,SQLERRM);
  END Import_Muster2;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
PROCEDURE import_auftrege(p_blob_id number, p_typ_id number, p_region_id number, p_out out varchar2) AS

    v_kennung               varchar2(10);
    v_auftrag_id            number;
    v_umzetzung_code        varchar2(20);
    v_datum                 date;
    v_check                 number;
    v_einlesung_status      varchar2(1) := 'Y';
    v_message               varchar2(4000);

    v_gaeb_version          VARCHAR2(20);
    v_item_count            NUMBER;

  BEGIN

    -- 1. PRE-CHECK: GAEB Version auslesen (für detaillierte Fehlermeldungen)
    BEGIN
        SELECT xt.version INTO v_gaeb_version
        FROM pd_import_x86 x,
             XMLTABLE(
                '/*[local-name()="GAEB"]/*[local-name()="GAEBInfo"]' 
                PASSING XMLTYPE(x.DATEI, nls_charset_id('AL32UTF8'))
                COLUMNS version VARCHAR2(20) PATH '*[local-name()="Version"]'
             ) xt
        WHERE x.id = p_blob_id;
    EXCEPTION WHEN OTHERS THEN
        v_gaeb_version := 'Unbekannt';
    END;

    -- 2. PRE-CHECK: Prüfen, ob überhaupt Items in der Datei existieren
    SELECT count(*) INTO v_item_count
    FROM pd_import_x86 x,
         XMLTABLE(
            '//*[local-name()="Item"]' 
            PASSING XMLTYPE(x.DATEI, nls_charset_id('AL32UTF8'))
         ) xt
    WHERE x.id = p_blob_id;

    IF v_item_count = 0 THEN
        raise_application_error(-20001, 
            'Fehler beim Einlesen: Es wurden keine Leistungspositionen gefunden. ' ||
            'Die Datei meldet sich als GAEB Version "' || v_gaeb_version || '". ' ||
            'Bitte prüfen Sie, ob es sich um eine gültige X83/X86 Datei handelt.'
        );
    END IF;

    -- =========================================================================
    -- SCHRITT 1: KOPFDATEN EINLESEN (Namespace unabhängig!)
    -- =========================================================================
    for i in (  
        SELECT  xt.name, xt.strasse, xt.postleitzahl, xt.Stadt, xt.Telefon, xt.Fax, xt.Name2, 
                xt.Projekt_name || '-' || xt.lv_name as projekt_name, xt.Projekt_desc, xt.sap_nr, xt.vertrag_nr,
                to_date(xt.Projekt_date, 'YYYY-MM-DD') Projekt_date, 
                -- Sichere Konvertierung der Total-Summe
                TO_NUMBER(REPLACE(TRIM(xt.total), '.', ',')) as total,
                -- Robustes Handling der Kreditorennummer (inkl. Format 3.3 Fallbacks)
                CASE
                  WHEN xt.KEDITOREN_NUMMER IS NULL OR TRIM(xt.KEDITOREN_NUMMER) IS NULL
                       OR NOT REGEXP_LIKE(xt.KEDITOREN_NUMMER, '^\s*[-+]?\d+([.,]\d+)?\s*$')
                    THEN '-'
                  ELSE TO_CHAR(TO_NUMBER(REPLACE(TRIM(xt.KEDITOREN_NUMMER), '.', ',')))
                END AS KEDITOREN_NUMMER
        FROM    pd_import_x86 x,
                XMLTABLE('/*[local-name()="GAEB"]'
                    PASSING XMLType(x.DATEI, nls_charset_id('AL32UTF8'))
                    COLUMNS
                        name          VARCHAR2(100)  PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="Name1"]',
                        name2         VARCHAR2(100)  PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="Name2"]',
                        strasse       VARCHAR2(100)  PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="Street"]',
                        postleitzahl  VARCHAR2(10)   PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="PCode"]',
                        Stadt         VARCHAR2(50)   PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="City"]',
                        Telefon       VARCHAR2(20)   PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="Phone"]',
                        Fax           VARCHAR2(20)   PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="Address"]/*[local-name()="Fax"]',
                        Projekt_name  VARCHAR2(100)  PATH '*[local-name()="PrjInfo"]/*[local-name()="NamePrj"]',
                        Projekt_desc  VARCHAR2(100)  PATH '*[local-name()="PrjInfo"]/*[local-name()="LblPrj"]',
                        Projekt_date  VARCHAR2(20)   PATH '*[local-name()="Award"]/*[local-name()="AwardInfo"]/*[local-name()="ContrDate"]',
                        SAP_nr        VARCHAR2(20)   PATH '*[local-name()="Award"]/*[local-name()="AwardInfo"]/*[local-name()="ContrNo"]',
                        vertrag_nr    VARCHAR2(20)   PATH '*[local-name()="Award"]/*[local-name()="OWN"]/*[local-name()="AwardNo"]',
                        total         VARCHAR2(30)   PATH '*[local-name()="Award"]/*[local-name()="BoQ"]/*[local-name()="BoQInfo"]/*[local-name()="Totals"]/*[local-name()="Total"]',
                        KEDITOREN_NUMMER VARCHAR2(50)PATH '*[local-name()="Award"]/*[local-name()="CTR"]/*[local-name()="AcctsPayNo"]',
                        lv_name       VARCHAR2(100)  PATH '*[local-name()="Award"]/*[local-name()="BoQ"]/*[local-name()="BoQInfo"]/*[local-name()="Name"]'
                ) xt
        where   x.id = p_blob_id
    ) loop
        begin
            select  distinct 1 into v_check
            from    PD_AUFTRAEGE
            where   sap_nr = i.sap_nr
            and     vertrag_nr = i.vertrag_nr;
        exception when no_data_found then
            v_check := 0;
        end;

        if v_check > 0  then
            v_einlesung_status := 'N';
            v_message := v_message || 'Der Vertrag mit ' || case when i.KEDITOREN_NUMMER = '-' then ' leerer Kreditoren Nummer, ' else null end ||
                                   ' SAP Nummer: ' || i.sap_nr || ' und Vertrag Nummer: ' || i.vertrag_nr || ' wurde schon eingelesen' || chr(10);
        elsif i.KEDITOREN_NUMMER = '-' or i.sap_nr is null or i.vertrag_nr is null then
            v_message := v_message || 'Der Vertrag mit ' || case when i.KEDITOREN_NUMMER = '-' then ' leerer Kreditoren Nummer, ' else null end ||
                                      case when i.sap_nr is null then ' leerer SAP Nummer ' else 'SAP Nummer: ' || i.sap_nr end ||
                                      case when i.vertrag_nr is null then 'und mit leerer Vertrag Nummer' else ' und Vertrag Nummer: ' || i.vertrag_nr end ||
                                      ' wurde erfolgreich eingelesen' || chr(10);
        else
            v_message := v_message || 'Der Vertrag mit SAP Nummer: ' || i.sap_nr || ' und Vertrag Nummer: ' || i.vertrag_nr || ' wurde erfolgreich eingelesen' || chr(10);
        end if;

        MERGE INTO PD_AUFTRAEGE e
        USING (SELECT i.name AUFTRAGNAHMER_NAME, i.Projekt_name PROJEKT_NAME FROM dual) h
        ON (lower(e.AUFTRAGNAHMER_NAME) = lower(h.AUFTRAGNAHMER_NAME) and lower(e.PROJEKT_NAME) = lower(h.PROJEKT_NAME))
        WHEN MATCHED THEN
        UPDATE SET e.strasse = nvl(i.strasse,strasse),
                   e.postleitzahl = nvl(i.postleitzahl,postleitzahl),
                   e.Stadt = nvl(i.Stadt,Stadt),
                   e.Telefon = nvl(i.Telefon,Telefon),
                   e.fax = nvl(i.fax,fax),
                   e.AUFTRAGNAHMER_NAME2 = nvl(i.name2,AUFTRAGNAHMER_NAME2),
                   e.PROJEKT_DESC = nvl(i.Projekt_desc,PROJEKT_DESC),
                   e.REGIONALBEREICH_ID = nvl(p_region_id,REGIONALBEREICH_ID),
                   e.datum = nvl(i.Projekt_date,datum),
                   e.total = nvl(i.total,total),
                   e.KEDITOREN_NUMMER = nvl(i.KEDITOREN_NUMMER,KEDITOREN_NUMMER),
                   e.sap_nr = nvl(i.sap_nr,sap_nr),
                   e.vertrag_nr = nvl(i.vertrag_nr,vertrag_nr),
                   e.DATUM_EINLESUNG = sysdate
        WHEN NOT MATCHED THEN
            INSERT (AUFTRAGNAHMER_NAME,strasse,postleitzahl,stadt,telefon,fax,AUFTRAGNAHMER_NAME2,PROJEKT_NAME,PROJEKT_DESC,REGIONALBEREICH_ID,DATUM,TOTAL,KEDITOREN_NUMMER,SAP_NR,VERTRAG_NR,EINLESUNG_STATUS,DATUM_EINLESUNG)
            VALUES (i.name, i.strasse,i.postleitzahl,i.stadt,i.telefon,i.fax,i.name2,i.Projekt_name,i.Projekt_desc,p_region_id,i.Projekt_date,i.total,i.KEDITOREN_NUMMER,i.sap_nr,i.vertrag_nr,v_einlesung_status,sysdate);

        select  id, datum into v_auftrag_id, v_datum
        from    PD_AUFTRAEGE
        where   lower(AUFTRAGNAHMER_NAME) = lower(i.name) and lower(PROJEKT_NAME) = lower(i.Projekt_name);
    end loop;


    -- =========================================================================
    -- SCHRITT 2: POSITIONEN EINLESEN (Namespace unabhängig!)
    -- =========================================================================
    for j in (  
        SELECT  xtd2.name2 as name,
                case 
                    when xtd2.name5 like 'MLV%' then trim(xtd2.name5)
                    when xtd2.name4 like 'MLV%' then trim(xtd2.name4)
                    when xtd2.name3 like 'MLV%' then trim(xtd2.name3)
                    when xtd2.name2 like '%MLV-%' or xtd2.name2 like '%MVL-%' then substr(replace(xtd2.name2,'MVL-','MLV-'),instr(replace(xtd2.name2,'MVL-','MLV-'),'MLV'),16)
                    else xt.code || '.' || xtd.code2 || '.' || xtd2.code3
                end as code,
                xtd2.description,
                -- Sicheres Fallback-Parsing analog Format 3.3
                to_number(replace(xtd2.menge,'.',',')) as menge,
                xtd2.ME,
                to_number(replace(xtd2.Einheitspreis,'.',',')) as Einheitspreis,
                to_number(replace(xtd2.Gesamtbetrag,'.',',')) as Gesamtbetrag,
                xt.code || '.' || xtd.code2 || '.' || xtd2.code3 as position
        FROM pd_import_x86 x,
             XMLTABLE('/*[local-name()="GAEB"]/*[local-name()="Award"]/*[local-name()="BoQ"]/*[local-name()="BoQBody"]/*[local-name()="BoQCtgy"]'
                        PASSING XMLType(x.DATEI,nls_charset_id('AL32UTF8'))
                        COLUMNS
                            code     VARCHAR2(2000)  PATH '@RNoPart',
                            name     VARCHAR2(4000)  PATH '*[local-name()="LblTx"]',
                            subitem  xmltype         PATH '*[local-name()="BoQBody"]/*[local-name()="BoQCtgy"]'
                     ) xt,
             xmltable('/*[local-name()="BoQCtgy"]'
                        PASSING xt.subitem
                        COLUMNS 
                            code2        VARCHAR2(2000)  PATH '@RNoPart',
                            name2        VARCHAR2(4000)  PATH '*[local-name()="LblTx"]',
                            subsubitem   XMLTYPE         PATH '*[local-name()="BoQBody"]/*[local-name()="Itemlist"]/*[local-name()="Item"]'
                     ) xtd,
             xmltable('/*[local-name()="Item"]'
                        PASSING xtd.subsubitem
                        COLUMNS 
                            code3            VARCHAR2(200)  PATH '@RNoPart',
                            name2            VARCHAR2(4000) PATH '*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="OutlineText"]/*[local-name()="OutlTxt"]/*[local-name()="TextOutlTxt"]/*[local-name()="p"]/*[local-name()="span"][1]',
                            name3            VARCHAR2(4000) PATH '*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="OutlineText"]/*[local-name()="OutlTxt"]/*[local-name()="TextOutlTxt"]/*[local-name()="p"]/*[local-name()="span"][2]',
                            name4            VARCHAR2(4000) PATH '*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="OutlineText"]/*[local-name()="OutlTxt"]/*[local-name()="TextOutlTxt"]/*[local-name()="p"]/*[local-name()="span"][3]',
                            name5            VARCHAR2(4000) PATH '*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="OutlineText"]/*[local-name()="OutlTxt"]/*[local-name()="TextOutlTxt"]/*[local-name()="p"]/*[local-name()="span"][4]',
                            description      VARCHAR2(4000) PATH 'substring(*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="DetailTxt"],1,4000)',
                            Menge            VARCHAR2(20)   PATH '*[local-name()="Qty"]',
                            ME               VARCHAR2(10)   PATH '*[local-name()="QU"]',
                            Einheitspreis    VARCHAR2(20)   PATH '*[local-name()="UP"]',
                            Gesamtbetrag     VARCHAR2(20)   PATH '*[local-name()="IT"]'
                     ) xtd2
        where x.id = p_blob_id
    ) loop
        begin
            select  NEW_KENNUNG into v_umzetzung_code
            from    pd_muster_umsetzung
            where   replace(replace(replace(OLD_KENNUNG,'_'),' '),'-') = replace(replace(replace(j.code,'_'),' '),'-');
        exception when no_data_found then
            v_umzetzung_code := null;
        end;

        begin
            select  1 into v_check
            from    PD_AUFTRAG_POSITIONEN
            where   auftrag_id = v_auftrag_id
            and     POSITION = j.position;
        exception when no_data_found then
            v_check := 0;
        end;

        if v_check = 0 then
            INSERT INTO PD_AUFTRAG_POSITIONEN(AUFTRAG_ID, NAME, CODE, BEZEICHNUNG, MENGE, MENGE_EINHEIT, EINHEITSPREIS, GESAMTBETRAG, UMZETZUNG_CODE, POSITION)
            VALUES (v_auftrag_id, j.name, j.code, j.description, j.menge, j.me, j.Einheitspreis, j.gesamtbetrag, nvl(v_umzetzung_code, j.code), j.position);
        end if;

        IF j.code like 'MLV%' then
            MERGE INTO PD_AUFTRAG_LVS al
            USING (Select substr(nvl(v_umzetzung_code, j.code), 1, 7) as code, v_auftrag_id as auftrag from dual) h
            ON (al.LV_CODE = h.code and al.AUFTRAG_ID = h.auftrag)
            WHEN NOT MATCHED THEN
                INSERT (LV_CODE, AUFTRAG_ID) VALUES (h.code, h.auftrag);
        END IF;
    end loop;

    delete from pd_import_x86 where id = p_blob_id;
    p_out := v_message;

EXCEPTION WHEN OTHERS THEN
    DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE, 'IMPORT');
    delete from pd_import_x86 where id = p_blob_id;
    raise_application_error(-20000, SQLERRM);
END import_auftrege;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE export_ausschreibung_to_excel(
    p_blob_id number, p_typ_id number, p_region_id number, 
    p_von date, p_bis date, p_regionen varchar2, 
    p_liferant varchar2, p_user_id number
) AS 
    -- Variablen für den XML Pre-Check
    v_item_count   NUMBER;
    v_gaeb_version VARCHAR2(20);

    -- Variablen für den Format-Check (Router)
    v_file_blob    BLOB;
    v_file_start   VARCHAR2(200);
    v_lob_length   NUMBER;
BEGIN
    DBS_LOGGING.LOG_INFO_AT ('Export_Ausschreibung', 'Start Routing & Pre-Check. BlobID: ' || p_blob_id);

    -- 1. FORMAT ERKENNEN: Die ersten Bytes lesen, OHNE es als XML zu parsen
    SELECT datei INTO v_file_blob FROM pd_import_x86 WHERE id = p_blob_id;
    v_lob_length := dbms_lob.getlength(v_file_blob);

    IF v_lob_length > 0 THEN
        -- Wandelt die ersten 100 Byte in Text um, um den Header zu lesen
        v_file_start := utl_raw.cast_to_varchar2(dbms_lob.substr(v_file_blob, 100, 1));
    ELSE
        raise_application_error(-20001, 'Die hochgeladene Datei ist leer (0 Bytes).');
    END IF;

    -- =========================================================================
    -- 2. ROUTER LOGIK
    -- =========================================================================

    IF v_file_start LIKE '%<?xml%' OR v_file_start LIKE '%<GAEB%' THEN
        DBS_LOGGING.LOG_INFO_AT ('Export_Ausschreibung', 'Format erkannt: GAEB XML');

        BEGIN
            SELECT xt.version INTO v_gaeb_version
            FROM pd_import_x86 x,
                 XMLTABLE(
                    '/*[local-name()="GAEB"]/*[local-name()="GAEBInfo"]' 
                    PASSING XMLTYPE(x.DATEI, nls_charset_id('AL32UTF8'))
                    COLUMNS version VARCHAR2(20) PATH '*[local-name()="Version"]'
                 ) xt
            WHERE x.id = p_blob_id;
        EXCEPTION WHEN OTHERS THEN
            v_gaeb_version := 'Unbekannt';
        END;

        SELECT count(*) INTO v_item_count
        FROM pd_import_x86 x,
             XMLTABLE(
                '//*[local-name()="Item"]' 
                PASSING XMLTYPE(x.DATEI, nls_charset_id('AL32UTF8'))
             ) xt
        WHERE x.id = p_blob_id;

        IF v_item_count = 0 THEN
            raise_application_error(-20001, 
                'Fehler beim Einlesen: Es wurden keine Leistungspositionen gefunden. ' ||
                'Die Datei meldet sich als GAEB Version "' || v_gaeb_version || '". ' ||
                'Möglicherweise ist die Datei leer oder weicht von der Standard GAEB-Struktur ab.'
            );
        END IF;

        export_ausschreibung_to_excel_unified(
            p_blob_id, p_typ_id, p_region_id, p_von, p_bis, p_regionen, p_liferant, p_user_id
        );

    ELSIF v_file_start LIKE '00        8%' THEN
        DBS_LOGGING.LOG_INFO_AT ('Export_Ausschreibung', 'Format erkannt: GAEB 90 (Text)');

        export_ausschreibung_to_excel_gaeb90(
            p_blob_id, p_typ_id, p_region_id, p_von, p_bis, p_regionen, p_liferant, p_user_id
        );

    ELSE
        -- Weder XML noch GAEB90
        raise_application_error(-20002, 'Das Dateiformat wird nicht unterstützt. Die Datei ist weder im GAEB XML- noch im GAEB 90-Format.');
    END IF;

-- =========================================================================
-- 4. GLOBALER FEHLERABFANG FÜR DEN E-MAIL VERSAND
-- =========================================================================
EXCEPTION 
    WHEN OTHERS THEN
        -- Fehler in die Tabelle loggen
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.export_ausschreibung_to_excel: Fehler: ' || SQLCODE || ': ' || SQLERRM ||
        ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSSCHREIBUNG');

        -- IMMER E-Mail an den User schicken, mit der genauen Fehlermeldung
        SendMailAuswertungFehler(
            p_user_id => p_user_id,
            p_error => '<b>Fehler beim Verarbeiten der Datei:</b><br><br>' || 
                       SQLERRM || '<br><br>' || 
                       '<i>Mögliche Ursachen: Das Format wird nicht unterstützt, die Struktur ist fehlerhaft oder keine Treffer zu den Eingabeparametern gefunden. ' ||
                       'Bitte verwenden Sie primär iTWO Exporte im GAEB XML Format und/oder Verwenden Sie keine Eingabeparameter.<br><br>' ||
                       'Wenn es trotzdem nicht geht, geben Sie die Datei bitte an unsere IT weiter, sodass das Problem schnellstmöglich behoben werden kann.</i>'
        );

        -- Fallback: Option zum Löschen der defekten Datei, damit kein Müll in der DB bleibt
        delete from pd_import_x86 where id = p_blob_id;

END export_ausschreibung_to_excel;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
PROCEDURE export_ausschreibung_to_excel_gaeb90(
    p_blob_id number, p_typ_id number, p_region_id number, p_von date, p_bis date, 
    p_regionen varchar2, p_liferant varchar2, p_user_id number
) AS 
       workbook xlsx_writer.book_r;
       sheet_1  integer;
       xlsx     blob;

       cs_border integer;
       cs_master integer;
       cs_master2 integer;
       cs_parent integer;
       number_format_child integer;
       number_format_parent integer;
       number_format_master integer;
       border_db_full integer;
       font_db  integer;
       fill_master integer;
       fill_parent integer;
       fill_master2 integer;

       l_min number; l_max number; l_avg number; l_median number; l_count number; l_id number;
       v_vergabesumme1 number; 
       v_vergabesumme2 number; 
       v_trimm number;
       l_trimm number;
       v_col_offset integer := 0;

       l_col              integer;
       v_has_menge        boolean;
       v_menge_check      number;

       c_menge_let        varchar2(5);
       c_dropdown_let     varchar2(5);
       c_val_ep_let       varchar2(5);
       c_val_gp_let       varchar2(5);
       c_mw_let           varchar2(5);
       c_med_let          varchar2(5);
       c_trim_let         varchar2(5);

       v_formula_ep       varchar2(4000);
       v_formula_gp       varchar2(4000);

       font_header_white      integer;
       fill_dark_blue_light   integer;
       border_thick           integer;
       cs_header              integer;

       l_row integer := 2;

       v_original_filename varchar2(600);
       v_excel_filename    varchar2(600);

       v_file_blob    BLOB;
       v_clob         CLOB;
       v_dest_offset  INTEGER := 1;
       v_src_offset   INTEGER := 1;
       v_lang_context INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       v_warning      INTEGER;

       v_line         VARCHAR2(32767);
       v_start_pos    INTEGER := 1;
       v_end_pos      INTEGER;

       TYPE t_hierarchy_map IS TABLE OF VARCHAR2(4000) INDEX BY VARCHAR2(100);
       v_hierarchies  t_hierarchy_map;
       v_current_lvl  VARCHAR2(100);

       v_pos_code     VARCHAR2(100);
       v_pos_menge    NUMBER;
       v_pos_me       VARCHAR2(20);
       v_pos_name     VARCHAR2(4000);
       v_pos_kennung  VARCHAR2(100);
       
       v_is_x82       NUMBER := 0;
       v_pos_ep       NUMBER := 0;
       v_pos_gp       NUMBER := 0;
       v_file_start   VARCHAR2(200);
       c_bepr_lv_let  VARCHAR2(5);
       c_abw_let      VARCHAR2(5);
       pct_format     INTEGER;
       v_formula_abw  VARCHAR2(4000);

       font_orange          integer;
       number_format_orange integer;

       PROCEDURE insert_parsed_position IS
           l_master_code VARCHAR2(10);
           l_parent_code VARCHAR2(10);
           l_item_code   VARCHAR2(10);
           l_master_name VARCHAR2(4000);
           l_parent_name VARCHAR2(4000);

           l_master_fmt  VARCHAR2(100);
           l_parent_fmt  VARCHAR2(100);
           l_code_fmt    VARCHAR2(100);
           l_clean_name  VARCHAR2(1000);
       BEGIN
           -- Split GAEB 90 code 
           l_master_code := SUBSTR(v_pos_code, 1, 2);
           l_parent_code := SUBSTR(v_pos_code, 1, 4);
           l_item_code   := SUBSTR(v_pos_code, 5, 4);

           IF v_hierarchies.exists(l_master_code) THEN l_master_name := v_hierarchies(l_master_code); END IF;
           IF v_hierarchies.exists(l_parent_code) THEN l_parent_name := v_hierarchies(l_parent_code); END IF;

           l_master_fmt := l_master_code;
           l_parent_fmt := l_master_code || '.' || SUBSTR(l_parent_code, 3, 2);
           l_code_fmt   := l_parent_fmt || '.' || l_item_code;

           l_clean_name := replace(substr(replace(v_pos_name,'MVL-','MLV-'),instr(replace(v_pos_name,'MVL-','MLV-'),'MLV-'),16),' ','_');
           IF instr(l_clean_name,'MLV') = 0 THEN 
               v_pos_kennung := '';
           ELSIF substr(l_clean_name,8,1) = '_' THEN 
               v_pos_kennung := l_clean_name;
           ELSE 
               v_pos_kennung := substr(l_clean_name,1,7) || '_' || substr(l_clean_name,8);
           END IF;

           BEGIN
                SELECT min(ap.EINHEITSPREIS), max(ap.EINHEITSPREIS), avg(ap.EINHEITSPREIS), median(ap.EINHEITSPREIS), count(ap.code)
                INTO l_min, l_max, l_avg, l_median, l_count
                FROM pd_auftrag_positionen ap
                JOIN pd_auftraege a ON a.id = ap.auftrag_id
                WHERE replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(v_pos_kennung,'-'),'_'),' ')
                AND a.EINLESUNG_STATUS = 'Y'
                AND a.REGIONALBEREICH_ID IN (SELECT column_value FROM table(apex_string.split_numbers(p_regionen, ':')))
                AND a.KEDITOREN_NUMMER IN (SELECT column_value FROM table(apex_string.split(p_liferant, ':')))
                AND ap.einheitspreis > 0
                AND (v_vergabesumme1 IS NULL OR a.total >= v_vergabesumme1)
                AND (v_vergabesumme2 IS NULL OR a.total <= v_vergabesumme2);

                IF v_trimm IS NOT NULL THEN
                    SELECT avg(case 
                                 when total_count <= 2 then avg_ep 
                                 when pct_rank >= (v_trimm / 100) and pct_rank <= (1 - (v_trimm / 100)) then avg_ep 
                                 else null 
                               end)
                    INTO l_trimm
                    FROM (
                        SELECT ap.einheitspreis as avg_ep, PERCENT_RANK() OVER (ORDER BY ap.einheitspreis) as pct_rank, COUNT(*) OVER () as total_count
                        FROM pd_auftrag_positionen ap
                        JOIN pd_auftraege a ON a.id = ap.auftrag_id
                        WHERE replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(v_pos_kennung,'-'),'_'),' ')
                        AND a.EINLESUNG_STATUS = 'Y'
                        AND a.REGIONALBEREICH_ID IN (SELECT column_value FROM table(apex_string.split_numbers(p_regionen, ':')))
                        AND a.KEDITOREN_NUMMER IN (SELECT column_value FROM table(apex_string.split(p_liferant, ':')))
                        AND ap.einheitspreis > 0
                        AND (v_vergabesumme1 IS NULL OR a.total >= v_vergabesumme1)
                        AND (v_vergabesumme2 IS NULL OR a.total <= v_vergabesumme2)
                    );
                ELSE
                    l_trimm := null;
                END IF;
           EXCEPTION
                WHEN no_data_found THEN 
                    l_min:=null; l_max:=null; l_avg:=null; l_median:=null; l_id :=null; l_trimm := null; l_count := 0;
           END;

           INSERT INTO pd_ausschreibung
                (blob_id, muster_id, name, code, einheit, menge, min_preis, mittlerer_preis, max_preis, median_preis, GETRIMMTER_PREIS, parent_id, master_id, parent_name, master_name, kennung, code_count, x82, einzelpreis, gesamtpreis)
           VALUES
                (p_blob_id, l_id, v_pos_name, l_code_fmt, v_pos_me, v_pos_menge, l_min, l_avg, l_max, l_median, l_trimm, l_parent_fmt, l_master_fmt, l_parent_name, l_master_name, v_pos_kennung, l_count, v_is_x82, v_pos_ep, v_pos_gp);
       END insert_parsed_position;

BEGIN
    DBS_LOGGING.LOG_INFO_AT ('Export_Ausschreibung', 'Start GAEB 90 Export. BlobID: ' || p_blob_id);

    BEGIN
        SELECT name, datei, GETRIMMTER_MITTELWERT, VERGABESUMME, VERGABESUMME2 
        INTO v_original_filename, v_file_blob, v_trimm, v_vergabesumme1, v_vergabesumme2 
        FROM pd_import_x86 
        WHERE id = p_blob_id;
        v_excel_filename := 'Ausschreibung_' || regexp_replace(v_original_filename, '\.[a-zA-Z0-9]+$', '') || '.xlsx';
    EXCEPTION WHEN NO_DATA_FOUND THEN
        v_excel_filename := 'Ausschreibung_LVS.xlsx';
        v_trimm := null; v_vergabesumme1 := null; v_vergabesumme2 := null;
    END;

    v_file_start := utl_raw.cast_to_varchar2(dbms_lob.substr(v_file_blob, 100, 1));
    IF lower(v_original_filename) LIKE '%.x82' OR 
       lower(v_original_filename) LIKE '%.p82' OR 
       lower(v_original_filename) LIKE '%.d82' OR
       v_file_start LIKE '00        82%'
    THEN
        v_is_x82 := 1;
    END IF;

    IF v_trimm IS NOT NULL THEN v_col_offset := 2; END IF;

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet(workbook, 'Ausschreibung LVS');

    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_master2 := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    font_header_white := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'theme="0"');
    fill_dark_blue_light := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="4D93D9"/><bgColor indexed="64"/></patternFill>');
    border_thick := xlsx_writer.add_border(workbook, '<left style="thick"><color indexed="64"/></left><right style="thick"><color indexed="64"/></right><top style="thick"><color indexed="64"/></top><bottom style="thick"><color indexed="64"/></bottom><diagonal/>');
    
    font_orange := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'rgb="D46F0A"');
    number_format_orange := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_orange);

    cs_header := xlsx_writer.add_cell_style(workbook, font_id => font_header_white, fill_id => fill_dark_blue_light, border_id => border_thick, vertical_alignment => 'center');
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, font_id => font_db);
    cs_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master,font_id => font_db);
    cs_master2 := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master2,font_id => font_db);
    cs_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_parent,font_id => font_db);
    number_format_child := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €",font_id => font_db);
    number_format_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db, fill_id => fill_parent);
    number_format_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db, fill_id => fill_master2);
    
    pct_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => 10, font_id => font_db);

    DBMS_LOB.CREATETEMPORARY(v_clob, TRUE);
    DBMS_LOB.CONVERTTOCLOB(v_clob, v_file_blob, DBMS_LOB.LOBMAXSIZE, v_dest_offset, v_src_offset, 
                           nls_charset_id('WE8MSWIN1252'), v_lang_context, v_warning);

    v_pos_code := NULL;
    LOOP
        v_end_pos := DBMS_LOB.INSTR(v_clob, CHR(10), v_start_pos);
        EXIT WHEN v_end_pos = 0 OR v_end_pos IS NULL;

        v_line := TRIM(REPLACE(DBMS_LOB.SUBSTR(v_clob, v_end_pos - v_start_pos, v_start_pos), CHR(13), ''));
        v_start_pos := v_end_pos + 1;

        IF SUBSTR(v_line, 1, 2) = '11' THEN
            v_current_lvl := TRIM(SUBSTR(v_line, 3, 9));
        ELSIF SUBSTR(v_line, 1, 2) = '12' THEN
            IF v_current_lvl IS NOT NULL THEN
                v_hierarchies(v_current_lvl) := TRIM(SUBSTR(v_line, 3, 72));
            END IF;
        ELSIF SUBSTR(v_line, 1, 2) = '21' THEN
            IF v_pos_code IS NOT NULL THEN
                insert_parsed_position;
            END IF;

            v_pos_code := TRIM(SUBSTR(v_line, 3, 9));
            BEGIN
                v_pos_menge := TO_NUMBER(TRIM(SUBSTR(v_line, 25, 11))) / 1000;
            EXCEPTION WHEN OTHERS THEN v_pos_menge := 0; END;
            
            v_pos_me := TRIM(SUBSTR(v_line, 36, 4));
            v_pos_name := '';
            
            v_pos_ep := 0;
            IF v_is_x82 = 1 THEN
                BEGIN
                    v_pos_ep := TO_NUMBER(TRIM(SUBSTR(v_line, 40, 9))) / 1000;
                EXCEPTION WHEN OTHERS THEN v_pos_ep := 0; END;
            END IF;
            v_pos_gp := v_pos_menge * v_pos_ep;

        ELSIF SUBSTR(v_line, 1, 2) = '25' THEN
            v_pos_name := v_pos_name || TRIM(SUBSTR(v_line, 3, 72));
        END IF;
    END LOOP;

    IF v_pos_code IS NOT NULL THEN
        insert_parsed_position;
    END IF;

    DBMS_LOB.FREETEMPORARY(v_clob);

    v_has_menge := false;
    begin
        select max(menge) into v_menge_check from pd_ausschreibung where blob_id = p_blob_id;
        if v_menge_check > 0 then
            v_has_menge := true;
        end if;
    exception when others then
        v_has_menge := false;
    end;

    l_col := 1;
    xlsx_writer.col_width(workbook, sheet_1, l_col, 40); l_col := l_col + 1; 
    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 
    xlsx_writer.col_width(workbook, sheet_1, l_col, 20); l_col := l_col + 1; 
    xlsx_writer.col_width(workbook, sheet_1, l_col, 10); l_col := l_col + 1; 
    xlsx_writer.col_width(workbook, sheet_1, l_col, 8);  l_col := l_col + 1; 
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- bepr. LV
    END IF;
    
    xlsx_writer.col_width(workbook, sheet_1, l_col, 20); l_col := l_col + 1; -- VAL. PREIS (Dropdown)
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 20); l_col := l_col + 1; -- Abw. %
    END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 
    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 19); l_col := l_col + 1; END IF;
    
    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 21); l_col := l_col + 1; END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 25); l_col := l_col + 1; END IF;

    IF v_trimm IS NOT NULL THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 18); l_col := l_col + 1; 
        IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 25); l_col := l_col + 1; END IF;
    END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; 
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 19); l_col := l_col + 1; END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 11); 
     
    l_col := 1;
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'NAME'); l_col := l_col + 1;
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'CODE'); l_col := l_col + 1;
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'KENNUNG'); l_col := l_col + 1;
    
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MENGE');
    c_menge_let := xlsx_writer.col_to_letter(l_col); 
    l_col := l_col + 1;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'EINHEIT'); l_col := l_col + 1;
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'bepr. LV');
        c_bepr_lv_let := xlsx_writer.col_to_letter(l_col);
        l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VALIDIERUNGSPREIS');
    c_dropdown_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'Abw. Planer/Val. In %');
        c_abw_let := xlsx_writer.col_to_letter(l_col);
        l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VAL. EP');
    c_val_ep_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VAL. GP'); 
    c_val_gp_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MIN PREIS'); l_col := l_col + 1;
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MIN PREIS'); l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MITTELWERT');
    c_mw_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MITTELWERT'); l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MEDIAN PREIS');
    c_med_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MEDIAN PREIS'); l_col := l_col + 1;
    END IF;

    IF v_trimm IS NOT NULL THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GETRIMMTER MW');
        c_trim_let := xlsx_writer.col_to_letter(l_col);
        l_col := l_col + 1;
        
        IF v_has_menge THEN
            xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT GETRIMMTER MW'); l_col := l_col + 1;
        END IF;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MAX PREIS'); l_col := l_col + 1;
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MAX PREIS'); l_col := l_col + 1;
    END IF;
    
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VERGABEN');

    -- NEW: Extracted x82 fields einzelpreis and gesamtpreis for all layers
    for j in (with  mustern as 
                   (select  pa.parent_id,pa.code,pa.kennung,pa.name,pa.min_preis,pa.mittlerer_preis,pa.max_preis,pa.median_preis,
                            pa.GETRIMMTER_PREIS,
                            pa.EINHEIT,pa.MENGE,'CHILD' art,
                            pa.min_preis*pa.MENGE gesamt_min,
                            pa.mittlerer_preis * pa.MENGE gesammt_mittel,
                            pa.max_preis * pa.MENGE gesamt_max,
                            pa.median_preis * pa.MENGE gesamt_median,
                            pa.GETRIMMTER_PREIS * pa.MENGE gesamt_trimm,
                            pa.code_count,
                            pa.einzelpreis,
                            pa.gesamtpreis
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    union all
                    select  pa.master_id,pa.parent_id,null,pa.parent_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),
                            sum(pa.GETRIMMTER_PREIS),
                            null,null,'PARENT' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),
                            sum(pa.GETRIMMTER_PREIS*pa.MENGE),
                            null,
                            null,
                            sum(pa.gesamtpreis)
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.parent_id,pa.master_id,pa.parent_name
                    union all 
                    select  null,pa.master_id,null,pa.master_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),
                            sum(pa.GETRIMMTER_PREIS),
                            null,null,'MASTER' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),
                            sum(pa.GETRIMMTER_PREIS*pa.MENGE),
                            null,
                            null,
                            sum(pa.gesamtpreis)
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.master_id,pa.master_name
                   )
                SELECT  trim(m.code) code,
                        trim(m.name) name,
                        trim(m.kennung) kennung,
                        trim(m.einheit) einheit,
                        m.menge,
                        m.min_preis,
                        m.mittlerer_preis,
                        m.max_preis,
                        m.median_preis,
                        m.getrimmter_preis,
                        m.gesamt_min,
                        m.gesammt_mittel,
                        m.gesamt_max,
                        m.gesamt_median,
                        m.gesamt_trimm,
                        m.art,
                        m.code_count,
                        m.einzelpreis,
                        m.gesamtpreis,
                        (case m.art
                            when 'CHILD' then cs_border
                            when 'PARENT' then cs_parent
                            else cs_master
                        end) cs_style,
                        (case m.art
                            when 'CHILD' then number_format_child
                            when 'PARENT' then number_format_parent
                            else number_format_master
                        end) number_format
              FROM    mustern m
              START WITH parent_id IS NULL
              CONNECT BY PRIOR m.code = m.parent_id 
              ORDER SIBLINGS BY m.code
             )
    loop
       l_col := 1;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.name); l_col := l_col + 1;
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.code); l_col := l_col + 1;
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.kennung); l_col := l_col + 1;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, value_ => j.menge); l_col := l_col + 1;
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.einheit); l_col := l_col + 1;

       IF j.art = 'CHILD' THEN
           IF v_is_x82 = 1 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.einzelpreis); 
               l_col := l_col + 1;
           END IF;

           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => 'Mittelwert'); l_col := l_col + 1;
           
           IF v_is_x82 = 1 THEN
               v_formula_abw := 'IF(OR(' || c_bepr_lv_let || l_row || '=0,' || c_val_ep_let || l_row || '=0),"",(' || c_val_ep_let || l_row || '-' || c_bepr_lv_let || l_row || ')/' || c_bepr_lv_let || l_row || ')';
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => pct_format, formula => v_formula_abw);
               l_col := l_col + 1;
           END IF;

           v_formula_ep := 'IF(' || c_dropdown_let || l_row || '="Mittelwert",' || c_mw_let || l_row || ',IF(' || c_dropdown_let || l_row || '="Median",' || c_med_let || l_row || ',IF(' || c_dropdown_let || l_row || '="getrimmter MW",' || case when v_trimm is not null then c_trim_let else 'ZZ' end || l_row || ',0)))';
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, formula => v_formula_ep); l_col := l_col + 1;
           
           v_formula_gp := c_menge_let || l_row || '*' || c_val_ep_let || l_row;
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, formula => v_formula_gp); l_col := l_col + 1;
       ELSE
           IF v_is_x82 = 1 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamtpreis); 
               l_col := l_col + 1;
           END IF;

           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
           
           IF v_is_x82 = 1 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
           END IF;

           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
       END IF;

       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.min_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_min); l_col := l_col + 1;
       END IF;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.mittlerer_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesammt_mittel); l_col := l_col + 1;
       END IF;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.median_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_median); l_col := l_col + 1;
       END IF;

       IF v_trimm IS NOT NULL THEN
           IF j.art = 'CHILD' AND j.code_count <= 2 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => number_format_orange, value_ => j.getrimmter_preis); l_col := l_col + 1;
               IF v_has_menge THEN
                   xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => number_format_orange, value_ => j.gesamt_trimm); l_col := l_col + 1;
               END IF;
           ELSE
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.getrimmter_preis); l_col := l_col + 1;
               IF v_has_menge THEN
                   xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_trimm); l_col := l_col + 1;
               END IF;
           END IF;
       END IF;

       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.max_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_max); l_col := l_col + 1;
       END IF;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, value_ => j.code_count);

       l_row := l_row + 1;
    end loop;

    IF l_row > 2 THEN
         xlsx_writer.add_data_validation(
            xlsx => workbook, 
            sheet => sheet_1, 
            sqref => c_dropdown_let || '2:' || c_dropdown_let || (l_row - 1), 
            formula1 => CASE 
                            WHEN v_trimm IS NOT NULL THEN '"Mittelwert,Median,getrimmter MW"'
                            ELSE '"Mittelwert,Median"'
                        END
        );
    END IF;

    xlsx_writer.freeze_sheet(workbook, sheet_1,0,1);
    workbook := print_param_worksheet(workbook, p_blob_id, number_format_child);
    xlsx := xlsx_writer.create_xlsx(workbook);

    delete from pd_import_x86 where id = p_blob_id;
    delete from pd_ausschreibung where blob_id = p_blob_id;

    SendMailAuswertung(p_user_id => p_user_id, p_anhang => xlsx, p_filename => v_excel_filename);

    DBMS_LOB.FREETEMPORARY(xlsx);

EXCEPTION 
    WHEN others THEN
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.export_ausschreibung_to_excel_gaeb90: Fehler bei Auswertung: ' || SQLCODE || ': ' || SQLERRM ||
        ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSSCHREIBUNG');

        RAISE_APPLICATION_ERROR(-20004, 'Fehler bei der GAEB 90 Verarbeitung: ' || SQLERRM);

END export_ausschreibung_to_excel_gaeb90;
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE export_ausschreibung_to_excel_unified(
    p_blob_id number, p_typ_id number, p_region_id number, p_von date, p_bis date, 
    p_regionen varchar2, p_liferant varchar2, p_user_id number
) AS 
       workbook xlsx_writer.book_r;
       sheet_1  integer;
       xlsx     blob;

       cs_border integer;
       cs_master integer;
       cs_master2 integer;
       cs_parent integer;
       number_format_child integer;
       number_format_parent integer;
       number_format_master integer;
       border_db_full integer;
       font_db  integer;
       fill_master integer;
       fill_parent integer;
       fill_master2 integer;

       l_min number; l_max number; l_avg number; l_median number; l_count number; l_id number;
       v_vergabesumme1 number; 
       v_vergabesumme2 number; 
       v_trimm number;
       l_trimm number;
       v_col_offset integer := 0;

       l_col              integer;
       v_has_menge        boolean;
       v_menge_check      number;

       c_menge_let        varchar2(5);
       c_dropdown_let     varchar2(5);
       c_val_ep_let       varchar2(5);
       c_val_gp_let       varchar2(5);
       c_mw_let           varchar2(5);
       c_med_let          varchar2(5);
       c_trim_let         varchar2(5);
       
       v_is_x82           number := 0;
       v_file_blob        blob;
       c_bepr_lv_let      varchar2(5);
       c_abw_let          varchar2(5);
       pct_format         integer;
       v_formula_abw      varchar2(4000);

       v_formula_ep       varchar2(4000);
       v_formula_gp       varchar2(4000);

       font_header_white      integer;
       fill_dark_blue_light   integer;
       border_thick           integer;
       cs_header              integer;
       
       l_row integer := 2;

       v_original_filename varchar2(600);
       v_excel_filename    varchar2(600);

       font_orange          integer;
       number_format_orange integer;
BEGIN

    DBS_LOGGING.LOG_INFO_AT ('Export_Ausschreibung', 'Start Unified Export. BlobID: ' || p_blob_id);

    BEGIN
        SELECT name, datei INTO v_original_filename, v_file_blob 
        FROM pd_import_x86 
        WHERE id = p_blob_id;

        v_original_filename := regexp_replace(v_original_filename, '\.[a-zA-Z0-9]+$', '');
        v_excel_filename := 'Ausschreibung_' || v_original_filename || '.xlsx';
    EXCEPTION 
        WHEN no_data_found THEN
            v_excel_filename := 'Ausschreibung_LVS.xlsx';
    END;

    IF lower(v_original_filename) LIKE '%.x82' OR 
       lower(v_original_filename) LIKE '%.p82' OR 
       lower(v_original_filename) LIKE '%.d82' OR
       dbms_lob.instr(v_file_blob, utl_raw.cast_to_raw('<DP>82</DP>')) > 0 OR
       dbms_lob.instr(v_file_blob, utl_raw.cast_to_raw('DP="82"')) > 0 
    THEN
        v_is_x82 := 1; 
    END IF;

    BEGIN
        select GETRIMMTER_MITTELWERT, VERGABESUMME, VERGABESUMME2 
        into v_trimm, v_vergabesumme1, v_vergabesumme2 
        from PD_IMPORT_X86 where id=p_blob_id;
    EXCEPTION WHEN NO_DATA_FOUND THEN
        v_trimm := null;
        v_vergabesumme1 := null;
        v_vergabesumme2 := null;
    END;

    IF v_trimm IS NOT NULL THEN
        v_col_offset := 2; 
    END IF;

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Ausschreibung LVS');

    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_master2 := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    font_header_white := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'theme="0"');
    fill_dark_blue_light := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="4D93D9"/><bgColor indexed="64"/></patternFill>');
    border_thick := xlsx_writer.add_border(workbook, '<left style="thick"><color indexed="64"/></left><right style="thick"><color indexed="64"/></right><top style="thick"><color indexed="64"/></top><bottom style="thick"><color indexed="64"/></bottom><diagonal/>');
    
    cs_header := xlsx_writer.add_cell_style(workbook, font_id => font_header_white, fill_id => fill_dark_blue_light, border_id => border_thick, vertical_alignment => 'center');
    
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, font_id => font_db);
    cs_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master,font_id => font_db);
    cs_master2 := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master2,font_id => font_db);
    cs_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_parent,font_id => font_db);
    number_format_child := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €",font_id => font_db);
    number_format_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db, fill_id => fill_parent);
    number_format_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db, fill_id => fill_master2);

    font_orange := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'rgb="D46F0A"');
    number_format_orange := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_orange);
    
    pct_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => 10, font_id => font_db);

    FOR i IN (
        SELECT 
            xt.item_name AS name,
            xt.master_code AS master,
            xt.master_code || '.' || xt.parent_code AS parent,
            xt.master_name AS master_name,
            xt.parent_name AS parent_name,

            (CASE
                WHEN instr(replace(substr(replace(xt.item_name,'MVL-','MLV-'),instr(replace(xt.item_name,'MVL-','MLV-'),'MLV-'),16),' ','_'),'MLV') = 0 THEN ''
                WHEN substr(replace(substr(replace(xt.item_name,'MVL-','MLV-'),instr(replace(xt.item_name,'MVL-','MLV-'),'MLV-'),16),' ','_'),8,1) = '_'
                     THEN replace(substr(replace(xt.item_name,'MVL-','MLV-'),instr(replace(xt.item_name,'MVL-','MLV-'),'MLV-'),16),' ','_')
                ELSE substr(replace(substr(replace(xt.item_name,'MVL-','MLV-'),instr(replace(xt.item_name,'MVL-','MLV-'),'MLV-'),16),' ','_'),1,7) || '_' || 
                     substr(replace(substr(replace(xt.item_name,'MVL-','MLV-'),instr(replace(xt.item_name,'MVL-','MLV-'),'MLV-'),16),' ','_'),8)                            
            END) AS kennung,

            xt.master_code || '.' || xt.parent_code || '.' || xt.item_code AS code,
            xt.description,
            replace(xt.menge, '.', ',') AS menge,
            xt.me,
            replace(xt.einheitspreis, '.', ',') AS einheitspreis,
            replace(xt.gesamtbetrag, '.', ',') AS gesamtbetrag
        FROM pd_import_x86 x,
                XMLTABLE(
                'for $i in //*[local-name()="Item"]
                 return <row>
                          {$i}
                          <item_name>{normalize-space(string-join($i/*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="OutlineText"]/*[local-name()="OutlTxt"]/*[local-name()="TextOutlTxt"]//text(), " "))}</item_name>
                          <master_code>{data($i/ancestor::*[local-name()="BoQCtgy"][1]/@RNoPart)}</master_code>
                          <master_name>{normalize-space($i/ancestor::*[local-name()="BoQCtgy"][1]/*[local-name()="LblTx"])}</master_name>
                          <parent_code>{data($i/ancestor::*[local-name()="BoQCtgy"][last()]/@RNoPart)}</parent_code>
                          <parent_name>{normalize-space($i/ancestor::*[local-name()="BoQCtgy"][last()]/*[local-name()="LblTx"])}</parent_name>
                        </row>'
                PASSING XMLType(x.DATEI, nls_charset_id('AL32UTF8'))
                COLUMNS
                    item_code     VARCHAR2(200)  PATH '*[local-name()="Item"]/@RNoPart',
                    item_name     VARCHAR2(4000) PATH 'item_name',
                    description   VARCHAR2(4000) PATH 'substring(*[local-name()="Item"]/*[local-name()="Description"]/*[local-name()="CompleteText"]/*[local-name()="DetailTxt"],1,4000)',
                    menge         VARCHAR2(20)   PATH '*[local-name()="Item"]/*[local-name()="Qty"]',
                    me            VARCHAR2(10)   PATH '*[local-name()="Item"]/*[local-name()="QU"]',
                    einheitspreis VARCHAR2(20)   PATH '*[local-name()="Item"]/*[local-name()="UP"]',
                    gesamtbetrag  VARCHAR2(20)   PATH '*[local-name()="Item"]/*[local-name()="IT"]',

                    master_code   VARCHAR2(200)  PATH 'master_code',
                    master_name   VARCHAR2(4000) PATH 'master_name',
                    parent_code   VARCHAR2(200)  PATH 'parent_code',
                    parent_name   VARCHAR2(4000) PATH 'parent_name'
             ) xt
        WHERE x.id = p_blob_id
    )
    LOOP
       BEGIN
            SELECT min(ap.EINHEITSPREIS), max(ap.EINHEITSPREIS), avg(ap.EINHEITSPREIS), median(ap.EINHEITSPREIS), count(ap.code)
            INTO l_min, l_max, l_avg, l_median, l_count
            FROM pd_auftrag_positionen ap
            JOIN pd_auftraege a ON a.id = ap.auftrag_id
            WHERE replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(i.kennung,'-'),'_'),' ')
            AND a.EINLESUNG_STATUS = 'Y'
            AND a.REGIONALBEREICH_ID IN (SELECT column_value FROM table(apex_string.split_numbers(p_regionen, ':')))
            AND a.KEDITOREN_NUMMER IN (SELECT column_value FROM table(apex_string.split(p_liferant, ':')))
            AND ap.einheitspreis > 0
            AND (v_vergabesumme1 IS NULL OR a.total >= v_vergabesumme1)
            AND (v_vergabesumme2 IS NULL OR a.total <= v_vergabesumme2);

            IF v_trimm IS NOT NULL THEN
                SELECT avg(case 
                             when total_count <= 2 then avg_ep 
                             when pct_rank >= (v_trimm / 100) and pct_rank <= (1 - (v_trimm / 100)) then avg_ep 
                             else null 
                           end)
                INTO l_trimm
                FROM (
                    SELECT ap.einheitspreis as avg_ep,
                           PERCENT_RANK() OVER (ORDER BY ap.einheitspreis) as pct_rank,
                           COUNT(*) OVER () as total_count
                    FROM pd_auftrag_positionen ap
                    JOIN pd_auftraege a ON a.id = ap.auftrag_id
                    WHERE replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(i.kennung,'-'),'_'),' ')
                    AND a.EINLESUNG_STATUS = 'Y'
                    AND a.REGIONALBEREICH_ID IN (SELECT column_value FROM table(apex_string.split_numbers(p_regionen, ':')))
                    AND a.KEDITOREN_NUMMER IN (SELECT column_value FROM table(apex_string.split(p_liferant, ':')))
                    AND ap.einheitspreis > 0
                    AND (v_vergabesumme1 IS NULL OR a.total >= v_vergabesumme1)
                    AND (v_vergabesumme2 IS NULL OR a.total <= v_vergabesumme2)
                );
            ELSE
                l_trimm := null;
            END IF;

       EXCEPTION
            WHEN no_data_found THEN 
                l_min:=null; l_max:=null; l_avg:=null; l_median:=null; l_id :=null; l_trimm := null;
       END;

       INSERT INTO pd_ausschreibung
            (blob_id, muster_id, name, code, einheit, menge, min_preis, mittlerer_preis, max_preis, median_preis, GETRIMMTER_PREIS, parent_id, master_id, parent_name, master_name, kennung, code_count, x82, einzelpreis, gesamtpreis)
       VALUES
            (p_blob_id, l_id, i.name, i.code, i.me, to_number(replace(nvl(i.menge,0),'.',',')), l_min, l_avg, l_max, l_median, l_trimm, i.parent, i.master, i.parent_name, i.master_name, i.kennung, l_count, v_is_x82, to_number(replace(nvl(i.einheitspreis,0),'.',',')), to_number(replace(nvl(i.einheitspreis,0),'.',','))*to_number(replace(nvl(i.menge,0),'.',',')));

    END LOOP;

    v_has_menge := false;
    begin
        select max(menge) into v_menge_check from pd_ausschreibung where blob_id = p_blob_id;
        if v_menge_check > 0 then
            v_has_menge := true;
        end if;
    exception when others then
        v_has_menge := false;
    end;

    -- 1. Dynamische Spaltenbreiten setzen 
    l_col := 1;
    xlsx_writer.col_width(workbook, sheet_1, l_col, 40); l_col := l_col + 1; -- NAME
    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- CODE
    xlsx_writer.col_width(workbook, sheet_1, l_col, 20); l_col := l_col + 1; -- KENNUNG
    xlsx_writer.col_width(workbook, sheet_1, l_col, 10); l_col := l_col + 1; -- MENGE
    xlsx_writer.col_width(workbook, sheet_1, l_col, 8);  l_col := l_col + 1; -- EINHEIT
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- bepr. LV
    END IF;
    
    xlsx_writer.col_width(workbook, sheet_1, l_col, 20); l_col := l_col + 1; -- VAL. PREIS (Dropdown)
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 20); l_col := l_col + 1; -- Abw. %
    END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- VAL. EP
    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- VAL. GP

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- MIN
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 19); l_col := l_col + 1; END IF;
    
    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- MW
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 21); l_col := l_col + 1; END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- MED
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 25); l_col := l_col + 1; END IF;

    IF v_trimm IS NOT NULL THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 18); l_col := l_col + 1; -- TRIMM
        IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 25); l_col := l_col + 1; END IF;
    END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15); l_col := l_col + 1; -- MAX
    IF v_has_menge THEN xlsx_writer.col_width(workbook, sheet_1, l_col, 19); l_col := l_col + 1; END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 11); -- VERGABEN
    
    -- 2. Dynamische Header setzen 
    l_col := 1;
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'NAME'); l_col := l_col + 1;
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'CODE'); l_col := l_col + 1;
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'KENNUNG'); l_col := l_col + 1;
    
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MENGE');
    c_menge_let := xlsx_writer.col_to_letter(l_col); 
    l_col := l_col + 1;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'EINHEIT'); l_col := l_col + 1;
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'bepr. LV');
        c_bepr_lv_let := xlsx_writer.col_to_letter(l_col);
        l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VALIDIERUNGSPREIS');
    c_dropdown_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    IF v_is_x82 = 1 THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'Abw. Planer/Val. In %');
        c_abw_let := xlsx_writer.col_to_letter(l_col);
        l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VAL. EP');
    c_val_ep_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VAL. GP'); 
    c_val_gp_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MIN PREIS'); l_col := l_col + 1;
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MIN PREIS'); l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MITTELWERT');
    c_mw_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MITTELWERT'); l_col := l_col + 1;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MEDIAN PREIS');
    c_med_let := xlsx_writer.col_to_letter(l_col);
    l_col := l_col + 1;
    
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MEDIAN PREIS'); l_col := l_col + 1;
    END IF;

    IF v_trimm IS NOT NULL THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GETRIMMTER MW');
        c_trim_let := xlsx_writer.col_to_letter(l_col);
        l_col := l_col + 1;
        
        IF v_has_menge THEN
            xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT GETRIMMTER MW'); l_col := l_col + 1;
        END IF;
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'MAX PREIS'); l_col := l_col + 1;
    IF v_has_menge THEN
        xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'GESAMT MAX PREIS'); l_col := l_col + 1;
    END IF;
    
    xlsx_writer.add_cell(workbook, sheet_1, 1, l_col, style_id => cs_header, text => 'VERGABEN');

    for j in (with  mustern as 
                   (select  pa.parent_id,pa.code,pa.kennung,pa.name,pa.min_preis,pa.mittlerer_preis,pa.max_preis,pa.median_preis,
                            pa.GETRIMMTER_PREIS,
                            pa.EINHEIT,pa.MENGE,'CHILD' art,
                            pa.min_preis*pa.MENGE gesamt_min,
                            pa.mittlerer_preis * pa.MENGE gesammt_mittel,
                            pa.max_preis * pa.MENGE gesamt_max,
                            pa.median_preis * pa.MENGE gesamt_median,
                            pa.GETRIMMTER_PREIS * pa.MENGE gesamt_trimm,
                            pa.code_count,
                            pa.einzelpreis,
                            pa.gesamtpreis
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    union all
                    select  pa.master_id,pa.parent_id,null,pa.parent_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),
                            sum(pa.GETRIMMTER_PREIS),
                            null,null,'PARENT' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),
                            sum(pa.GETRIMMTER_PREIS*pa.MENGE),
                            null,
                            null,
                            sum(pa.gesamtpreis)
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.parent_id,pa.master_id,pa.parent_name
                    union all 
                    select  null,pa.master_id,null,pa.master_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),
                            sum(pa.GETRIMMTER_PREIS),
                            null,null,'MASTER' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),
                            sum(pa.GETRIMMTER_PREIS*pa.MENGE),
                            null,
                            null,
                            sum(pa.gesamtpreis)
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.master_id,pa.master_name
                   )
                SELECT  trim(m.code) code,
                        trim(m.name) name,
                        trim(m.kennung) kennung,
                        trim(m.einheit) einheit,
                        m.menge,
                        m.min_preis,
                        m.mittlerer_preis,
                        m.max_preis,
                        m.median_preis,
                        m.getrimmter_preis,
                        m.gesamt_min,
                        m.gesammt_mittel,
                        m.gesamt_max,
                        m.gesamt_median,
                        m.gesamt_trimm,
                        m.art,
                        m.code_count,
                        m.einzelpreis,
                        m.gesamtpreis,
                        (case m.art
                            when 'CHILD' then cs_border
                            when 'PARENT' then cs_parent
                            else cs_master
                        end) cs_style,
                        (case m.art
                            when 'CHILD' then number_format_child
                            when 'PARENT' then number_format_parent
                            else number_format_master
                        end) number_format
              FROM    mustern m
              START WITH parent_id IS NULL
              CONNECT BY PRIOR m.code = m.parent_id 
              ORDER SIBLINGS BY m.code
             )
    loop
       l_col := 1;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.name); l_col := l_col + 1;
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.code); l_col := l_col + 1;
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.kennung); l_col := l_col + 1;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, value_ => j.menge); l_col := l_col + 1;
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => j.einheit); l_col := l_col + 1;

       IF j.art = 'CHILD' THEN
       
           IF v_is_x82 = 1 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.einzelpreis); 
               l_col := l_col + 1;
           END IF;
       
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => 'Mittelwert'); l_col := l_col + 1;
        
           IF v_is_x82 = 1 THEN
               v_formula_abw := 'IF(OR(' || c_bepr_lv_let || l_row || '=0,' || c_val_ep_let || l_row || '=0),"",(' || c_val_ep_let || l_row || '-' || c_bepr_lv_let || l_row || ')/' || c_bepr_lv_let || l_row || ')';
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => pct_format, formula => v_formula_abw);
               l_col := l_col + 1;
           END IF;
           
           v_formula_ep := 'IF(' || c_dropdown_let || l_row || '="Mittelwert",' || c_mw_let || l_row || ',IF(' || 
           c_dropdown_let || l_row || '="Median",' || c_med_let || l_row || ',IF(' || c_dropdown_let || l_row || '="getrimmter MW",' || case when v_trimm is not null then c_trim_let else 'ZZ' end || l_row || ',0)))';
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, formula => v_formula_ep); l_col := l_col + 1;
           
           v_formula_gp := c_menge_let || l_row || '*' || c_val_ep_let || l_row;
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, formula => v_formula_gp); l_col := l_col + 1;
       ELSE
           IF v_is_x82 = 1 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamtpreis); 
               l_col := l_col + 1;
           END IF;
           
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
           
           IF v_is_x82 = 1 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
           END IF;
           
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, text => null); l_col := l_col + 1;
       END IF;

       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.min_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_min); l_col := l_col + 1;
       END IF;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.mittlerer_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesammt_mittel); l_col := l_col + 1;
       END IF;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.median_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_median); l_col := l_col + 1;
       END IF;

       IF v_trimm IS NOT NULL THEN
           IF j.art = 'CHILD' AND j.code_count <= 2 THEN
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => number_format_orange, value_ => j.getrimmter_preis); l_col := l_col + 1;
               IF v_has_menge THEN
                   xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => number_format_orange, value_ => j.gesamt_trimm); l_col := l_col + 1;
               END IF;
           ELSE
               xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.getrimmter_preis); l_col := l_col + 1;
               IF v_has_menge THEN
                   xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_trimm); l_col := l_col + 1;
               END IF;
           END IF;
       END IF;

       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.max_preis); l_col := l_col + 1;
       IF v_has_menge THEN
           xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.number_format, value_ => j.gesamt_max); l_col := l_col + 1;
       END IF;
       
       xlsx_writer.add_cell(workbook, sheet_1, l_row, l_col, style_id => j.cs_style, value_ => j.code_count);

       l_row := l_row + 1;
    end loop;

    IF l_row > 2 THEN
         xlsx_writer.add_data_validation(
            xlsx     => workbook, 
            sheet    => sheet_1, 
            sqref    => c_dropdown_let || '2:' || c_dropdown_let || (l_row - 1), 
            formula1 => CASE 
                            WHEN v_trimm IS NOT NULL THEN '"Mittelwert,Median,getrimmter MW"'
                            ELSE '"Mittelwert,Median"'
                        END
        );
    END IF;

    xlsx_writer.freeze_sheet(workbook, sheet_1,0,1);

    workbook := print_param_worksheet(workbook, p_blob_id, number_format_child);

    xlsx := xlsx_writer.create_xlsx(workbook);

    delete from pd_import_x86 where id = p_blob_id;
    delete from pd_ausschreibung where blob_id = p_blob_id;

    SendMailAuswertung(p_user_id => p_user_id, p_anhang => xlsx, p_filename => v_excel_filename);

    DBMS_LOB.FREETEMPORARY(xlsx);

EXCEPTION 
    WHEN others THEN
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.export_ausschreibung_to_excel_unified: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
        ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSSCHREIBUNG');

        SendMailAuswertungFehler(
            p_user_id => p_user_id,
            p_error => 'Fehler bei der Auswertung, bitte wenden Sie sich an den Anwendungsverantwortlichen.'
        );

END export_ausschreibung_to_excel_unified;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

procedure auswertung_to_excel_2(p_von date,
                                p_bis date,
                                p_lvs varchar2,
                                p_regionen varchar2,
                                p_liferant varchar2,
                                p_trimm number,   
                                p_vergabesumme_von number default null,
                                p_vergabesumme_bis number default null,
                                p_user_id number) as

       workbook xlsx_writer.book_r;
       sheet_1  integer;

       xlsx     blob;

       c_x_split constant integer := 3;
       c_y_split constant integer := 14;
       c_y_region constant integer := 1;

       -- Styles Variablen
       cs_center_wrapped integer;
       cs_center integer;
       cs_center_bold integer;
       cs_center_bold_white integer;
       cs_center_bold_grey integer;
       cs_border integer;
       cs_master_parent integer;
       cs_master_master integer;
       cs_rot integer;
       cs_gelb integer;
       cs_hellgelb integer;
       cs_hellgruen integer;
       cs_gruen integer;

       fill_gelb integer;
       fill_hellgelb integer;
       fill_hellgruen integer;
       fill_gruen integer;
       fill_rot integer;
       fill_db integer;
       fill_db_grey integer;
       fill_master integer;
       fill_parent integer;

       font_db  integer;
       font_db_small integer;
       font_db_bold integer;
       font_db_bold_white integer;
       border_db integer;
       border_db_full integer;

       number_format integer;
       pct_format    integer;
       center_number_format integer;
       datum_format integer;
       blob_id number ;

       font_orange          integer;
       number_format_orange integer;

       type t_contract_rec is record(
            region            varchar2(100),
            bezeichnung       varchar2(4000),
            auftragnahmer     varchar2(4000),
            lvdatum           date,
            vergabesumme      number,
            auftrag_id        number,
            kreditorennummer  varchar2(400),
            sap_nr            varchar2(400),
            vertrag_nr        varchar2(400),
            anzahl            varchar2(200),
            prozentual_anzahl varchar2(200),
            lv_code           varchar2(4000),
            column_position   number
       );
       type t_contract_tab is table of t_contract_rec;

       type t_muster_rec is record(
            parent_master     varchar2(10),
            position          varchar2(200),
            pos_text          varchar2(4000),
            einheit           varchar2(50),
            menge_soll        number,
            minimum_preis     number,
            mittelwert_preis  number,
            median_preis      number,
            getrimmter_preis  number,
            maximum_preis     number,
            anzahl            number,
            anzahl_getrimmt   number,   
            position_kennung  varchar2(400),
            sanitized_kennung varchar2(400),
            row_position      number
       );
       type t_muster_tab is table of t_muster_rec;

       type t_price_rec is record(
            auftrag_id       number,
            position_kennung varchar2(400),
            einheitspreis    number
       );
       type t_price_tab is table of t_price_rec;

       type t_price_map is table of number index by varchar2(800);

       l_contracts        t_contract_tab;
       l_mustern          t_muster_tab;
       l_price_rows       t_price_tab;
       l_price_map        t_price_map;

       l_column           number := 1;
       l_contract_count   number;
       l_row_index        number;
       l_col_index        number;
       l_style            integer;
       l_count_style      integer;
       l_price_key        varchar2(800);

       l_blob             blob;

       l_col                  integer;
       l_fixed_cols           integer;
       l_col_getrimmt         integer;
       l_contracts_start_col  integer;

       fill_black       integer;
       border_black     integer;
       cs_legende_header integer;
       cs_leg_rot       integer;
       cs_leg_gelb      integer;
       cs_leg_hellgelb  integer;
       cs_leg_hellgruen integer;
       cs_leg_gruen     integer;

begin

    DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.auswertung_to_excel_2 START: p_von='||to_char(p_von,'YYYY-MM-DD')||', p_bis='||to_char(p_bis,'YYYY-MM-DD')||', p_lvs='||nvl(p_lvs,'')||', p_regionen='||nvl(p_regionen,'')||', p_liferant='||nvl(p_liferant,'')||', p_user_id='||nvl(to_char(p_user_id),'-'),'AUSWERTUNG');

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Auswertung Vertrags LVS');

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    -- Fonts
    font_db := xlsx_writer.add_font(workbook, 'DB Office', 10);
    font_db_small := xlsx_writer.add_font(workbook, 'DB Office', 7);
    font_db_bold := xlsx_writer.add_font(workbook, 'DB Office', 10, b => true);
    font_db_bold_white := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'theme="0"', b => true);

    -- Borders
    border_db := xlsx_writer.add_border(workbook, '<left/><right/><top/><bottom/><diagonal/>');
    border_db_full := xlsx_writer.add_border(workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>' );
    border_black := xlsx_writer.add_border(workbook, '<left style="medium"><color rgb="000000"/></left><right style="medium"><color rgb="000000"/></right><top style="medium"><color rgb="000000"/></top><bottom style="medium"><color rgb="000000"/></bottom><diagonal/>');

    -- Fills
    fill_db := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="00ccff"/><bgColor indexed="64"/></patternFill>');
    fill_db_grey := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="d9d9d9"/><bgColor indexed="64"/></patternFill>');
    fill_master := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    fill_gelb := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="ffc000"/><bgColor indexed="64"/></patternFill>');
    fill_hellgelb := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="ffff00"/><bgColor indexed="64"/></patternFill>');
    fill_hellgruen := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="00ff00"/><bgColor indexed="64"/></patternFill>');
    fill_gruen := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="00b050"/><bgColor indexed="64"/></patternFill>');
    fill_rot := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="ff0000"/><bgColor indexed="64"/></patternFill>');
    fill_black := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="000000"/><bgColor indexed="64"/></patternFill>');
 
    -- Spezielle Styles nur für die Legende
    cs_legende_header := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_black, border_id => border_black, vertical_alignment => 'center');
    cs_leg_rot := xlsx_writer.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', font_id => font_db, border_id => border_black);
    cs_leg_gelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', font_id => font_db, border_id => border_black);
    cs_leg_hellgelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', font_id => font_db, border_id => border_black);
    cs_leg_hellgruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', font_id => font_db, border_id => border_black);
    cs_leg_gruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', font_id => font_db, border_id => border_black);

    cs_center_wrapped := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', wrap_text => true, font_id => font_db_small, border_id => border_db_full);
    cs_center := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full);
    cs_center_bold := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_db, border_id => border_db_full);
    cs_center_bold_white := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_db, vertical_alignment => 'center', border_id => border_db_full);
    cs_center_bold_grey := xlsx_writer.add_cell_style(workbook, fill_id => fill_db_grey, border_id => border_db_full, font_id => font_db_bold);
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full);
    cs_master_parent := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
    cs_master_master := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);

    cs_rot := xlsx_writer.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full);
    cs_gelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full);
    cs_hellgelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full);
    cs_hellgruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full);
    cs_gruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full);

    number_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db);
    center_number_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db, vertical_alignment => 'center');

    datum_format := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', font_id => font_db, border_id => border_db_full, num_fmt_id => xlsx_writer."mm-dd-yy");
    pct_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => 10, font_id => font_db);

    font_orange := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'rgb="D46F0A"');
    number_format_orange := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_orange);

    -- ==========================================
    -- I.) GRUNDANGABEN (Zeilen 1 - 13)
    -- ==========================================
    xlsx_writer.add_cell(workbook, sheet_1,  1, 1, style_id => cs_center_bold, text => 'I.) Grundangaben');

    xlsx_writer.add_cell(workbook, sheet_1,  2, 1, style_id => cs_border, text => 'Region');
    xlsx_writer.add_cell(workbook, sheet_1,  3, 1, style_id => cs_border, text => 'Bezeichnung der Maßnahme');
    xlsx_writer.add_cell(workbook, sheet_1,  4, 1, style_id => cs_border, text => 'Auftragnehmer');
    xlsx_writer.add_cell(workbook, sheet_1,  5, 1, style_id => cs_border, text => 'LV-Datum');
    xlsx_writer.add_cell(workbook, sheet_1,  6, 1, style_id => cs_border, text => 'Vergabesumme');
    xlsx_writer.add_cell(workbook, sheet_1,  7, 1, style_id => cs_border, text => 'Vergabevorgangsnummer');
    xlsx_writer.add_cell(workbook, sheet_1,  8, 1, style_id => cs_border, text => 'SAP-Kontraktnummer');
    xlsx_writer.add_cell(workbook, sheet_1,  9, 1, style_id => cs_border, text => 'Kreditorennummer');
    xlsx_writer.add_cell(workbook, sheet_1, 10, 1, style_id => cs_border, text => 'Anzahl Muster-LV Pos. mit Kennung');
    xlsx_writer.add_cell(workbook, sheet_1, 11, 1, style_id => cs_border, text => '% Anteil  Muster-LV Pos. mit Kennung');
    xlsx_writer.add_cell(workbook, sheet_1, 12, 1, style_id => cs_border, text => 'Muster-LV Code');

    -- Zeile 13: Überschrift Teil II
    xlsx_writer.add_cell(workbook, sheet_1, 13, 1, style_id => cs_center_bold, text => 'II.) Einheitspreise');

    l_col := 1;
    xlsx_writer.col_width(workbook, sheet_1, l_col, 40);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Pos.'); l_col := l_col + 1;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 50);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Pos.-Text'); l_col := l_col + 1;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 10);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Einheit'); l_col := l_col + 1;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Minimum'); l_col := l_col + 1;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Mittelwert'); l_col := l_col + 1;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Median'); l_col := l_col + 1;

    IF p_trimm IS NOT NULL THEN
        xlsx_writer.col_width(workbook, sheet_1, l_col, 15);
        xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Getrimmter MW'); l_col := l_col + 1;
    END IF;

    xlsx_writer.col_width(workbook, sheet_1, l_col, 15);
    xlsx_writer.add_cell(workbook, sheet_1, 14, l_col, style_id => cs_center_bold_grey, text => 'Maximum'); l_col := l_col + 1;

    l_fixed_cols := l_col;
    xlsx_writer.col_width(workbook, sheet_1, l_fixed_cols, 12);
    l_col := l_col + 1;

    IF p_trimm IS NOT NULL THEN
        l_col_getrimmt := l_col;
        xlsx_writer.col_width(workbook, sheet_1, l_col_getrimmt, 11);
        l_col := l_col + 1;
    END IF;

    l_contracts_start_col := l_col;

    -- Daten laden
    select  r.code region,
            a.projekt_desc bezeichnung,
            a.auftragnahmer_name auftragnahmer,
            a.datum lvdatum,
            a.total vergabesumme,
            a.id auftrag_id,
            a.keditoren_nummer,
            a.sap_nr,
            a.vertrag_nr,
            nvl(kennung_anzahl,0)||' von '||gesamt_anzahl anzahl,
            round((nvl(kennung_anzahl,0)/gesamt_anzahl)*100,2)||'%' prozentual_anzahl,
            listagg(lvs.lv_code, ', ') within group (order by lvs.lv_code) lv_code,
            null column_position
    bulk collect into l_contracts
    from    pd_auftraege a
            join (select count(*) gesamt_anzahl, auftrag_id from pd_auftrag_positionen group by auftrag_id) all_positionen
                on all_positionen.auftrag_id = a.id
            join (select count(*) kennung_anzahl, auftrag_id
                  from pd_auftrag_positionen
                  where umsetzung_code like 'MLV%'
                  group by auftrag_id) lv_positionen
                on lv_positionen.auftrag_id = a.id
            join pd_region r on r.id = a.regionalbereich_id
            join pd_auftrag_lvs lvs on lvs.auftrag_id = a.id
    where   a.datum between p_von and p_bis
    and     a.einlesung_status = 'Y'
    and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
    and     a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
    and     (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
    and     (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
    and     (
                select count(*)
                from   pd_auftrag_positionen ap
                       cross join pd_muster_lvs m
                where  m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
                and    ap.auftrag_id = a.id
                and    ap.code like 'M%'
                and    ap.einheitspreis > 0
                and    instr(ap.umsetzung_code, m.position_kennung2) > 0
            ) > 0
    group by r.code, a.projekt_desc, a.auftragnahmer_name, a.datum, a.total, a.id, a.keditoren_nummer, a.sap_nr, a.vertrag_nr, kennung_anzahl, gesamt_anzahl
    order by a.auftragnahmer_name;

    if l_contracts.count > 0 then
        for idx in 1..l_contracts.count loop
            l_col_index := l_contracts_start_col + l_column - 1;
            l_contracts(idx).column_position := l_col_index;

            xlsx_writer.col_width(workbook, sheet_1, l_col_index, 15);
            xlsx_writer.add_cell(workbook, sheet_1, 14, l_col_index, style_id => cs_center_bold_grey, text => '');

            -- Zeile 2 bis 12
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 1, l_col_index, style_id => cs_center, text => trim(l_contracts(idx).region));
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 2, l_col_index, style_id => cs_center_wrapped, text => trim(l_contracts(idx).bezeichnung));
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 3, l_col_index, style_id => cs_center, text => trim(l_contracts(idx).auftragnahmer));
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 4, l_col_index, style_id => cs_center, text => l_contracts(idx).lvdatum);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 5, l_col_index, style_id => center_number_format, value_ => l_contracts(idx).vergabesumme);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 6, l_col_index, style_id => cs_center, text => l_contracts(idx).vertrag_nr);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 7, l_col_index, style_id => cs_center, text => l_contracts(idx).sap_nr);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 8, l_col_index, style_id => cs_center, text => l_contracts(idx).kreditorennummer);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 9, l_col_index, style_id => cs_center, text => l_contracts(idx).anzahl);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 10, l_col_index, style_id => cs_center, text => l_contracts(idx).prozentual_anzahl);
            xlsx_writer.add_cell(workbook, sheet_1, c_y_region + 11, l_col_index, style_id => cs_center, text => l_contracts(idx).lv_code);

            l_column := l_column + 1;
        end loop;
    end if;

    l_contract_count := l_column - 1;

    xlsx_writer.add_cell(workbook, sheet_1, 14, l_fixed_cols, style_id => cs_center_bold_grey, text => l_contract_count || ' Vergaben');
    IF p_trimm IS NOT NULL THEN
        xlsx_writer.add_cell(workbook, sheet_1, 14, l_col_getrimmt, style_id => cs_center_bold_grey, text => 'Getrimmt');
    END IF;

    xlsx_writer.add_cell(workbook, sheet_1, 7, l_fixed_cols, style_id => cs_legende_header, text => 'Legende:');
    xlsx_writer.add_cell(workbook, sheet_1, 8, l_fixed_cols, style_id => cs_leg_rot, text => '= 0');
    xlsx_writer.add_cell(workbook, sheet_1, 9, l_fixed_cols, style_id => cs_leg_gelb, text => '= 1 - 4');
    xlsx_writer.add_cell(workbook, sheet_1, 10, l_fixed_cols, style_id => cs_leg_hellgelb, text => '= 5 - 9');
    xlsx_writer.add_cell(workbook, sheet_1, 11, l_fixed_cols, style_id => cs_leg_hellgruen, text => '= 10 - 24');
    xlsx_writer.add_cell(workbook, sheet_1, 12, l_fixed_cols, style_id => cs_leg_gruen, text => '≥ 25');

    select parent_master,
            position,
            pos_text,
            einheit,
            1, 
            minimum_preis,
            mittelwert_preis,
            median_preis,
            getrimmter_preis,
            maximum_preis,
            anzahl,
            anzahl_getrimmt,
            position_kennung,
            sanitized_kennung,
            null row_position
    bulk collect into l_mustern
    from (
            with mustern as (
                select  m.id,
                        m.code,
                        m.position_kennung,
                        m.position_kennung2,
                        m.name,
                        m.muster_typ_id || m.code id_tree,
                        decode(m.parent_id, null, null, m.muster_typ_id || m.parent_id) parent_tree,
                        m.parent_id,
                        m.einheit,
                        round(min(ap.einheitspreis), 2) minimum_preis,
                        round(avg(ap.einheitspreis), 2) mittelwert_preis,
                        round(median(ap.einheitspreis), 2) median_preis,
                        round(max(ap_trimm.trimm_preis), 2) getrimmter_preis, 
                        round(max(ap.einheitspreis), 2) maximum_preis,
                        count(ap.umsetzung_code) anzahl,
                        max(ap_trimm.anzahl_getrimmt) anzahl_getrimmt 
                from    pd_muster_lvs m
                left join (
                            select avg(pa.einheitspreis) einheitspreis,
                                    pa.auftrag_id,
                                    pa.umsetzung_code
                            from   pd_auftrag_positionen pa
                                    join pd_auftraege a on a.id = pa.auftrag_id
                                    join pd_region r on r.id = a.regionalbereich_id
                            where  a.datum between p_von and p_bis
                            and    a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
                            and    pa.code like 'M%'
                            and    a.einlesung_status = 'Y'
                            and    pa.einheitspreis > 0
                            and    r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                            and    (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
                            and    (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
                            group by pa.auftrag_id, pa.umsetzung_code
                          ) ap on m.position_kennung2 = ap.umsetzung_code
                left join (
                            select umsetzung_code,
                                   avg(case 
                                         when total_count <= 2 then avg_ep 
                                         when pct_rank >= (nvl(p_trimm, 0) / 100) 
                                          and pct_rank <= (1 - (nvl(p_trimm, 0) / 100)) 
                                         then avg_ep 
                                         else null 
                                       end) as trimm_preis,
                                   count(case 
                                         when total_count <= 2 then avg_ep 
                                         when pct_rank >= (nvl(p_trimm, 0) / 100) 
                                          and pct_rank <= (1 - (nvl(p_trimm, 0) / 100)) 
                                         then avg_ep 
                                         else null 
                                       end) as anzahl_getrimmt
                            from (
                                select pa.umsetzung_code,
                                       avg(pa.einheitspreis) as avg_ep,
                                       PERCENT_RANK() OVER (PARTITION BY pa.umsetzung_code ORDER BY avg(pa.einheitspreis)) as pct_rank,
                                       COUNT(*) OVER (PARTITION BY pa.umsetzung_code) as total_count
                                from pd_auftrag_positionen pa
                                join pd_auftraege a on a.id = pa.auftrag_id
                                join pd_region r on r.id = a.regionalbereich_id
                                where a.datum between p_von and p_bis
                                and a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
                                and pa.code like 'M%'
                                and a.einlesung_status = 'Y'
                                and pa.einheitspreis > 0
                                and r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                                and (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
                                and (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
                                group by pa.auftrag_id, pa.umsetzung_code
                            )
                            group by umsetzung_code
                          ) ap_trimm on m.position_kennung2 = ap_trimm.umsetzung_code
                group by m.id, m.code, m.name, m.position_kennung, m.position_kennung2, m.muster_typ_id, m.parent_id, m.einheit, m.muster_typ_id || m.code, decode(m.parent_id, null, null, m.muster_typ_id || m.parent_id)
            )
            select case
                        when m.parent_tree is null then 'MASTER'
                        when m.parent_id = '01' then 'PARENT'
                        else 'CHILD'
                   end parent_master,
                   m.code position,
                   m.name pos_text,
                   m.einheit,
                   case when m.parent_tree is null then null else m.minimum_preis end minimum_preis,
                   case when m.parent_tree is null then null else m.mittelwert_preis end mittelwert_preis,
                   case when m.parent_tree is null then null else m.median_preis end median_preis,
                   case when m.parent_tree is null then null else m.getrimmter_preis end getrimmter_preis,
                   case when m.parent_tree is null then null else m.maximum_preis end maximum_preis,
                   case when m.parent_tree is null or m.parent_id = '01' then null else m.anzahl end anzahl,
                   case when m.parent_tree is null or m.parent_id = '01' then null else m.anzahl_getrimmt end anzahl_getrimmt,
                   case when m.parent_tree is null then 'MASTER'
                        when m.parent_id = '01' then 'PARENT'
                        else m.position_kennung end position_kennung,
                   case when m.parent_tree is null then 'MASTER'
                        when m.parent_id = '01' then 'PARENT'
                        else replace(m.position_kennung2, ' ', '') end sanitized_kennung
            from mustern m
            start with m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
            connect by prior id_tree = parent_tree
            order siblings by m.code
         );

    if l_mustern.count > 0 then
        for idx in 1..l_mustern.count loop
            l_row_index := c_y_split + idx;

            if l_mustern(idx).parent_master = 'PARENT' then
                l_style := cs_master_parent;
            elsif l_mustern(idx).parent_master = 'MASTER' then
                l_style := cs_master_master;
            else
                l_style := cs_border;
            end if;

            l_col := 1;
            xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => trim(l_mustern(idx).position)); l_col := l_col + 1;
            xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => trim(l_mustern(idx).pos_text)); l_col := l_col + 1;
            xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => trim(l_mustern(idx).einheit)); l_col := l_col + 1;

            if l_mustern(idx).parent_master = 'PARENT' then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => 'Minimum'); l_col := l_col + 1;
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => 'Mittelwert'); l_col := l_col + 1;
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => 'Median'); l_col := l_col + 1;
                IF p_trimm IS NOT NULL THEN
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => 'Getrimmter MW'); l_col := l_col + 1;
                END IF;
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => 'Maximum'); l_col := l_col + 1;
            elsif l_mustern(idx).parent_master = 'MASTER' then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => ''); l_col := l_col + 1;
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => ''); l_col := l_col + 1;
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => ''); l_col := l_col + 1;
                IF p_trimm IS NOT NULL THEN
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => ''); l_col := l_col + 1;
                END IF;
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => ''); l_col := l_col + 1;
            else
                if l_mustern(idx).minimum_preis is not null then xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => number_format, value_ => l_mustern(idx).minimum_preis); end if;
                l_col := l_col + 1;
                if l_mustern(idx).mittelwert_preis is not null then xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => number_format, value_ => l_mustern(idx).mittelwert_preis); end if;
                l_col := l_col + 1;
                if l_mustern(idx).median_preis is not null then xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => number_format, value_ => l_mustern(idx).median_preis); end if;
                l_col := l_col + 1;
                IF p_trimm IS NOT NULL THEN
                    if l_mustern(idx).getrimmter_preis is not null then 
                        IF l_mustern(idx).anzahl <= 2 THEN
                            xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => number_format_orange, value_ => l_mustern(idx).getrimmter_preis);
                        ELSE
                            xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => number_format, value_ => l_mustern(idx).getrimmter_preis);
                        END IF;
                    end if;
                    l_col := l_col + 1;
                END IF;
                if l_mustern(idx).maximum_preis is not null then xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => number_format, value_ => l_mustern(idx).maximum_preis); end if;
                l_col := l_col + 1;
            end if;

        -- Anzahl Vergaben (Gesamt)
            if l_mustern(idx).parent_master in ('MASTER', 'PARENT') then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => '');
            elsif nvl(l_mustern(idx).anzahl, 0) = 0 then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_rot, value_ => 0);
            elsif l_mustern(idx).anzahl between 1 and 4 then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_gelb, value_ => l_mustern(idx).anzahl);
            elsif l_mustern(idx).anzahl between 5 and 9 then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_hellgelb, value_ => l_mustern(idx).anzahl);
            elsif l_mustern(idx).anzahl between 10 and 24 then
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_hellgruen, value_ => l_mustern(idx).anzahl);
            else
                xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_gruen, value_ => l_mustern(idx).anzahl);
            end if;
            l_col := l_col + 1;

            -- Anzahl Vergaben (Getrimmt)
            IF p_trimm IS NOT NULL THEN
                if l_mustern(idx).parent_master in ('MASTER', 'PARENT') then
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => l_style, text => '');
                elsif nvl(l_mustern(idx).anzahl_getrimmt, 0) = 0 then
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_rot, value_ => 0);
                elsif l_mustern(idx).anzahl_getrimmt between 1 and 4 then
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_gelb, value_ => l_mustern(idx).anzahl_getrimmt);
                elsif l_mustern(idx).anzahl_getrimmt between 5 and 9 then
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_hellgelb, value_ => l_mustern(idx).anzahl_getrimmt);
                elsif l_mustern(idx).anzahl_getrimmt between 10 and 24 then
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_hellgruen, value_ => l_mustern(idx).anzahl_getrimmt);
                else
                    xlsx_writer.add_cell(workbook, sheet_1, l_row_index, l_col, style_id => cs_gruen, value_ => l_mustern(idx).anzahl_getrimmt);
                end if;
                l_col := l_col + 1;
            END IF;

            l_mustern(idx).row_position := l_row_index;
        end loop;
    end if;

    select auftrag_id,
           umsetzung_code position_kennung,
           round(avg(einheitspreis), 2) einheitspreis
    bulk collect into l_price_rows
    from pd_auftrag_positionen pa
         join pd_auftraege a on a.id = pa.auftrag_id
         join pd_region r on r.id = a.regionalbereich_id
    where a.datum between p_von and p_bis
    and   a.einlesung_status = 'Y'
    and   a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
    and   r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
    and   pa.einheitspreis > 0
    and   pa.code like 'M%'
    and   (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
    and   (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
    group by auftrag_id, umsetzung_code;

    if l_price_rows.count > 0 then
        for idx in 1..l_price_rows.count loop
            if l_price_rows(idx).position_kennung is not null then
                l_price_key := l_price_rows(idx).auftrag_id || '|' || replace(replace(replace(l_price_rows(idx).position_kennung, '-'), '_'), ' ', '');
                l_price_map(l_price_key) := l_price_rows(idx).einheitspreis;
            end if;
        end loop;
    end if;

    if l_contracts.count > 0 and l_mustern.count > 0 then
        for idx in 1..l_contracts.count loop
            for jdx in 1..l_mustern.count loop
                if l_mustern(jdx).parent_master = 'PARENT' then
                    xlsx_writer.add_cell(workbook, sheet_1, l_mustern(jdx).row_position, l_contracts(idx).column_position, style_id => cs_master_parent, text => '');
                elsif l_mustern(jdx).parent_master = 'MASTER' then
                    xlsx_writer.add_cell(workbook, sheet_1, l_mustern(jdx).row_position, l_contracts(idx).column_position, style_id => cs_master_master, text => '');
                else
                    l_price_key := l_contracts(idx).auftrag_id || '|' || l_mustern(jdx).sanitized_kennung;
                    if l_price_map.exists(l_price_key) then
                        xlsx_writer.add_cell(workbook, sheet_1, l_mustern(jdx).row_position, l_contracts(idx).column_position, style_id => number_format, value_ => l_price_map(l_price_key));
                    end if;
                end if;
            end loop;
        end loop;
    end if;

    workbook := print_param_worksheet(workbook, p_von, p_bis, p_regionen, p_lvs, p_liferant, p_trimm, p_vergabesumme_von, p_vergabesumme_bis);

    xlsx_writer.freeze_sheet(workbook, sheet_1, l_contracts_start_col - 1, c_y_split);

    xlsx := xlsx_writer.create_xlsx(workbook);

    SendMailAuswertung(p_user_id => p_user_id, p_anhang => xlsx, p_filename => 'Auswertung.xlsx');

    if dbms_lob.istemporary(xlsx) = 1 then
        dbms_lob.freetemporary(xlsx);
    end if;
    if dbms_lob.istemporary(l_blob) = 1 then
        dbms_lob.freetemporary(l_blob);
    end if;

exception
    when others then
        if dbms_lob.istemporary(xlsx) = 1 then
            dbms_lob.freetemporary(xlsx);
        end if;
        if dbms_lob.istemporary(l_blob) = 1 then
            dbms_lob.freetemporary(l_blob);
        end if;

        dbs_logging.log_error_at('PREISDATENBANK_PKG.auswertung_to_excel_2: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
            ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSWERTUNG');

        SendMailAuswertungFehler(
            p_user_id => p_user_id,
            p_error   => 'PREISDATENBANK_PKG.auswertung_to_excel_2: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM || ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE
        );
end auswertung_to_excel_2;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

procedure auswertung_to_excel_leser(p_von date, 
                                    p_bis date, 
                                    p_lvs varchar2,
                                    p_regionen varchar2,
                                    p_liferant varchar2,
                                    p_trimm number default null, 
                                    p_vergabesumme_von number default null,
                                    p_vergabesumme_bis number default null,
                                    p_user_id number) as

       workbook xlsx_writer.book_r;
       sheet_1  integer;

       xlsx     blob;

       c_limit   constant integer := 50;
       c_x_split constant integer := 3;
       c_y_split constant integer := 1;
       c_y_region constant integer := 1;

       cs_center_wrapped integer;
       cs_center integer;
       cs_number integer;
       cs_center_bold integer;
       cs_center_bold_white integer;
       cs_center_bold_grey integer;
       cs_border integer;
       cs_master integer;
       cs_master_master integer;
       cs_master_parent integer;
       cs_rot integer;
       cs_gelb integer;
       cs_hellgelb integer;
       cs_hellgruen integer;
       cs_gruen integer;
       cs_noborder integer;

       fill_gelb integer;
       fill_hellgelb integer;
       fill_hellgruen integer;
       fill_gruen integer;
       fill_rot integer;

       font_db  integer;
       font_db_small integer;
       font_db_bold integer;
       fill_db integer;
       fill_db_grey integer;
       font_db_bold_white integer;
       fill_master integer;
       fill_parent integer;
       border_db integer;
       border_db_full integer;

       -- Variablen für die braune Legende
       fill_braun       integer;
       border_braun     integer;
       cs_legende_header integer;
       cs_leg_rot       integer;
       cs_leg_gelb      integer;
       cs_leg_hellgelb  integer;
       cs_leg_hellgruen integer;
       cs_leg_gruen     integer;

       number_format integer;
       center_number_format integer;
       datum_format integer;

       TYPE curtype IS REF CURSOR;

       v_auftrage t_auftraege := new t_auftraege();
       v_kennungen T_KENNUNG := new T_KENNUNG();
       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;
       l_column         number:=1;
       l_row            number;
       l_betrag         number;
       l_random_number  number := floor(dbms_random.value(1, 1000000000));
begin

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Auswertung Vertrags LVS');

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    -- Style definition
    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    border_db := xlsx_writer.add_border  (workbook, '<left/><right/><top/><bottom/><diagonal/>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    fill_db:= xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00ccff"/><bgColor indexed="64"/></patternFill>');
    fill_db_grey := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="d9d9d9"/><bgColor indexed="64"/></patternFill>');
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    fill_gelb := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ffc000"/><bgColor indexed="64"/></patternFill>');
    fill_hellgelb := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ffff00"/><bgColor indexed="64"/></patternFill>');
    fill_hellgruen := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00ff00"/><bgColor indexed="64"/></patternFill>');
    fill_gruen := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00b050"/><bgColor indexed="64"/></patternFill>');
    fill_rot := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ff0000"/><bgColor indexed="64"/></patternFill>');

    font_db_small := xlsx_writer.add_font      (workbook, 'DB Office', 7);
    font_db_bold := xlsx_writer.add_font      (workbook, 'DB Office', 10, b => true);
    font_db_bold_white := xlsx_writer.add_font(workbook, 'DB Office', 10, color=> 'theme="0"', b => true); -- KORRIGIERT: Weiße Schrift

    -- Braune Legende definieren
    fill_braun := xlsx_writer.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="8B4513"/><bgColor indexed="64"/></patternFill>');
    border_braun := xlsx_writer.add_border(workbook, '<left style="medium"><color rgb="8B4513"/></left><right style="medium"><color rgb="8B4513"/></right><top style="medium"><color rgb="8B4513"/></top><bottom style="medium"><color rgb="8B4513"/></bottom><diagonal/>');

    cs_legende_header := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_braun, border_id => border_braun, vertical_alignment => 'center', vertical_horizontal => 'center');
    cs_leg_rot := xlsx_writer.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_braun);
    cs_leg_gelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_braun);
    cs_leg_hellgelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_braun);
    cs_leg_hellgruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_braun);
    cs_leg_gruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_braun);

    cs_center_wrapped  := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', wrap_text => true,font_id => font_db_small, border_id => border_db_full);
    cs_center  := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_center_bold  := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_db, border_id => border_db_full);
    cs_center_bold_white := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_db,vertical_alignment => 'center', vertical_horizontal => 'center', border_id => border_db_full);
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full);
    cs_noborder := xlsx_writer.add_cell_style(workbook, border_id => border_db);
    cs_center_bold_grey := xlsx_writer.add_cell_style(workbook, fill_id => fill_db_grey, border_id => border_db_full, font_id => font_db_bold);

    -- Standard Farben für Daten-Zellen
    cs_rot := xlsx_writer.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_gelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_hellgelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_hellgruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_gruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);

    cs_master_parent := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
    cs_master_master := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);

    number_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €", font_id => font_db);
    center_number_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00 €" , font_id => font_db, vertical_alignment => 'center', vertical_horizontal => 'center');
    datum_format := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full, num_fmt_id => xlsx_writer."mm-dd-yy");

    -- Wir zählen hier lediglich die passenden Verträge hoch, ohne diese ins Excel zu drucken
    for i in 
    (
        with all_positionen as (select count(*) gesamt_anzahl, AUFTRAG_ID from pd_auftrag_positionen group by AUFTRAG_ID),
             lv_positionen as (select count(*) kennung_anzahl, AUFTRAG_ID from pd_auftrag_positionen where UMZETZUNG_CODE like 'MLV%' group by AUFTRAG_ID)
        select  a.id
        from    pd_auftraege a
        join    pd_region r on r.id = a.regionalbereich_id
        join    lv_positionen ap on a.id = ap.auftrag_id
        join    pd_auftrag_lvs lvs on a.id = lvs.auftrag_id
        join    all_positionen allp on a.id = allp.AUFTRAG_ID
        where   a.datum between p_von and p_bis
        and     a.EINLESUNG_STATUS = 'Y'
        and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
        and     a.KEDITOREN_NUMMER in (select column_value from table(apex_string.split(p_liferant, ':')))
        and     (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
        and     (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
        and     (   select  count(*)
                    from        pd_auftrag_positionen ap
                    cross join  pd_muster_lvs m
                    where   m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
                    and     ap.auftrag_id = a.id
                    and     ap.code like 'M%'
                    and     ap.einheitspreis > 0
                    and     (instr(ap.UMSETZUNG_CODE,m.position_kennung2) > 0)
                ) > 0
        group by a.id
    )
    loop
        l_column:=l_column+1;
    end loop;

    -- ==========================================
    -- HEADER
    -- ==========================================
    xlsx_writer.add_cell(workbook, sheet_1, 1, 1, style_id => cs_center_bold_grey, text => 'Pos.');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 2, style_id => cs_center_bold_grey, text => 'Pos.-Text');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 3, style_id => cs_center_bold_grey, text => 'Einheit');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 4, style_id => cs_center_bold_grey, text => 'Minimum');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 5, style_id => cs_center_bold_grey, text => 'Mittelwert');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 6, style_id => cs_center_bold_grey, text => 'Median');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 7, style_id => cs_center_bold_grey, text => 'Getrimmter MW');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 8, style_id => cs_center_bold_grey, text => 'Maximum');

    xlsx_writer.add_cell(workbook, sheet_1, 1, 9, style_id => cs_center_bold_grey, text => (l_column-1) || ' Vergaben');

    xlsx_writer.add_cell(workbook, sheet_1, 2, 11, style_id => cs_legende_header, text => 'Legende:');
    xlsx_writer.add_cell(workbook, sheet_1, 3, 11, style_id => cs_leg_rot, text => '= 0');
    xlsx_writer.add_cell(workbook, sheet_1, 4, 11, style_id => cs_leg_gelb, text => '= 1 - 4');
    xlsx_writer.add_cell(workbook, sheet_1, 5, 11, style_id => cs_leg_hellgelb, text => '= 5 - 9');
    xlsx_writer.add_cell(workbook, sheet_1, 6, 11, style_id => cs_leg_hellgruen, text => '= 10 - 24');
    xlsx_writer.add_cell(workbook, sheet_1, 7, 11, style_id => cs_leg_gruen, text => '≥ 25');

    -- Spaltenbreiten
    xlsx_writer.col_width(workbook, sheet_1, 1, 20);
    xlsx_writer.col_width(workbook, sheet_1, 2, 80);
    xlsx_writer.col_width(workbook, sheet_1, 3, 10);
    xlsx_writer.col_width(workbook, sheet_1, 4, 15);
    xlsx_writer.col_width(workbook, sheet_1, 5, 15);
    xlsx_writer.col_width(workbook, sheet_1, 6, 15);
    xlsx_writer.col_width(workbook, sheet_1, 7, 20);
    xlsx_writer.col_width(workbook, sheet_1, 8, 15);
    xlsx_writer.col_width(workbook, sheet_1, 9, 15);
    xlsx_writer.col_width(workbook, sheet_1, 10, 10);
    xlsx_writer.col_width(workbook, sheet_1, 11, 15);

    l_row:=1;

    for j in (
              with  mustern as (select  m.id,
                                        m.code,
                                        m.position_kennung,
                                        m.name,
                                        m.description,
                                        m.MUSTER_TYP_ID || m.code id_tree, 
                                        decode(m.parent_id,null,null,m.MUSTER_TYP_ID||m.parent_id) parent_tree,
                                        m.parent_id,
                                        m.einheit,
                                        -- Umwandlung der Aggregation in formatierte Strings
                                        to_char(round(min(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MINIMUM_PREIS,
                                        to_char(round(avg(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MITTELWERT_PREIS, 
                                        to_char(round(median(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MEDIAN_PREIS,
                                        -- Getrimmter Mittelwert
                                        to_char(round(max(ap_trimm.trimm_preis),2),'999G999G999G999G990D00') GETRIMMTER_PREIS,
                                        to_char(round(max(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MAXIMUM_PREIS,
                                        COUNT(ap.UMSETZUNG_CODE) ANZAHL
                              from PD_MUSTER_LVS m
                              left join (select avg(EINHEITSPREIS) EINHEITSPREIS,auftrag_id,UMSETZUNG_CODE 
                                          from pd_auftrag_positionen pa
                                          join pd_auftraege a on a.id = pa.auftrag_id and a.datum between p_von and p_bis
                                          and a.KEDITOREN_NUMMER in (select column_value  
                                                      from table(apex_string.split(p_liferant, ':')))
                                          and  pa.code like 'M%'
                                          and  a.EINLESUNG_STATUS = 'Y'
                                         left join pd_region r on r.id = a.regionalbereich_id
                                         where EINHEITSPREIS > 0
                                         and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                                         and     (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
                                         and     (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
                                         group by auftrag_id,UMSETZUNG_CODE) ap on
                                            (m.position_kennung2 = ap.UMSETZUNG_CODE)
                              -- Left join für Getrimmten Wert
                              left join (
                                    select umsetzung_code,
                                           avg(case 
                                                 when total_count <= 2 then avg_ep 
                                                 when pct_rank >= (nvl(p_trimm, 0) / 100) 
                                                  and pct_rank <= (1 - (nvl(p_trimm, 0) / 100)) 
                                                 then avg_ep 
                                                 else null 
                                               end) as trimm_preis
                                    from (
                                        select pa.umsetzung_code,
                                               avg(pa.einheitspreis) as avg_ep,
                                               PERCENT_RANK() OVER (PARTITION BY pa.umsetzung_code ORDER BY avg(pa.einheitspreis)) as pct_rank,
                                               COUNT(*) OVER (PARTITION BY pa.umsetzung_code) as total_count
                                        from pd_auftrag_positionen pa
                                        join pd_auftraege a on a.id = pa.auftrag_id
                                        join pd_region r on r.id = a.regionalbereich_id
                                        where a.datum between p_von and p_bis
                                        and a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
                                        and pa.code like 'M%'
                                        and a.einlesung_status = 'Y'
                                        and pa.einheitspreis > 0
                                        and r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                                        and     (p_vergabesumme_von is null or a.total >= p_vergabesumme_von)
                                        and     (p_vergabesumme_bis is null or a.total <= p_vergabesumme_bis)
                                        group by pa.auftrag_id, pa.umsetzung_code
                                    )
                                    group by umsetzung_code
                                ) ap_trimm on (m.position_kennung2 = ap_trimm.umsetzung_code)
                              group by m.id,m.position_kennung,m.code,m.name,m.description,m.MUSTER_TYP_ID,m.einheit,m.parent_id)
                                        SELECT case when m.parent_tree is null then 'MASTER'
                                                    when m.parent_id = '01' then 'PARENT'
                                                    else 'CHILD' end PARENT_MASTER,
                                               m.code POSITION,
                                               m.name as POS_TEXT,
                                               m.EINHEIT,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Minimum'
                                                    else m.MINIMUM_PREIS end MINIMUM_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Mittelwert'
                                                    else m.MITTELWERT_PREIS end MITTELWERT_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Median'
                                                    else m.MEDIAN_PREIS end MEDIAN_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Getrimmter MW'
                                                    else m.GETRIMMTER_PREIS end GETRIMMTER_PREIS, 
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Maximum'
                                                    else m.MAXIMUM_PREIS end MAXIMUM_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then null
                                                    else m.ANZAHL end ANZAHL,
                                               case when m.parent_tree is null then 'MASTER'
                                                    when m.parent_id = '01' then 'PARENT'
                                                    else m.position_kennung end position_kennung
                                        FROM mustern m
                                        START WITH m.id in (select column_value  
                                                      from table(apex_string.split_numbers(p_lvs, ':')))
                                        CONNECT BY PRIOR id_tree = parent_tree
                                        ORDER SIBLINGS BY m.code
    ) loop

                        IF j.PARENT_MASTER = 'PARENT' then
                            cs_master:=xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
                        ELSIF j.PARENT_MASTER = 'MASTER' then
                            cs_master:=xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);
                        ELSE
                            cs_master:=cs_border;
                        END IF;

                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 1,style_id => cs_master, text => trim(j.POSITION));
                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 2,style_id => cs_master, text => trim(j.POS_TEXT));
                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 3,style_id => cs_master, text => trim(j.EINHEIT));

                        -- KORREKTUR für ORA-06502: Sichere to_number() Rück-Konvertierung!
                        if upper(j.MINIMUM_PREIS)='MINIMUM' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 4,style_id => cs_master, text => trim(j.MINIMUM_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 4,style_id => number_format, value_ => to_number(trim(j.MINIMUM_PREIS), '999G999G999G999G990D00'));
                        end if;

                        if upper(j.MITTELWERT_PREIS)='MITTELWERT' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 5,style_id => cs_master, text => trim(j.MITTELWERT_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 5,style_id => number_format, value_ => to_number(trim(j.MITTELWERT_PREIS), '999G999G999G999G990D00'));
                        end if;

                        if upper(j.MEDIAN_PREIS)='MEDIAN' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 6,style_id => cs_master, text => trim(j.MEDIAN_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 6,style_id => number_format, value_ => to_number(trim(j.MEDIAN_PREIS), '999G999G999G999G990D00'));
                        end if;

                        if upper(j.GETRIMMTER_PREIS)='GETRIMMTER MW' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 7,style_id => cs_master, text => trim(j.GETRIMMTER_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 7,style_id => number_format, value_ => to_number(trim(j.GETRIMMTER_PREIS), '999G999G999G999G990D00'));
                        end if;

                        if upper(j.MAXIMUM_PREIS)='MAXIMUM' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 8,style_id => cs_master, text => trim(j.MAXIMUM_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 8,style_id => number_format, value_ => to_number(trim(j.MAXIMUM_PREIS), '999G999G999G999G990D00'));
                        end if;

                        if j.ANZAHL = 0 then
                            cs_master := cs_rot;
                        elsif j.ANZAHL > 0 and j.ANZAHL < 5 then
                            cs_master := cs_gelb;
                        elsif j.ANZAHL > 4 and j.ANZAHL < 10 then
                            cs_master := cs_hellgelb;
                        elsif j.ANZAHL > 9 and j.ANZAHL < 25 then
                            cs_master := cs_hellgruen;
                        elsif j.ANZAHL > 24 then
                            cs_master := cs_gruen;
                        end if;
                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 9,style_id => cs_master, value_ => j.ANZAHL);

                        v_kennungen.extend;
                        v_kennungen(v_kennungen.count).kennung := j.position_kennung;
                        v_kennungen(v_kennungen.count).row_position := l_row+c_y_split;

                        l_row:=l_row+1;
    end loop;

    workbook := print_param_worksheet(workbook, p_von, p_bis, p_regionen, p_lvs, p_liferant, p_trimm, p_vergabesumme_von, p_vergabesumme_bis);

    xlsx_writer.freeze_sheet(workbook, sheet_1, c_x_split, c_y_split);
    xlsx     := xlsx_writer.create_xlsx(workbook);

    -- Mailversand
    SendMailAuswertung(p_user_id,p_anhang => xlsx,p_filename => 'Auswertung_Leser.xlsx');

    DBMS_LOB.FREETEMPORARY(xlsx);

exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.auswertung_to_excel_leser: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
      ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSWERTUNG');

        SendMailAuswertungFehler
            (
            p_user_id => p_user_id,
            p_error => 'PREISDATENBANK_PKG.auswertung_to_excel_leser: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE
            );

end auswertung_to_excel_leser;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function print_param_worksheet(p_ws  IN OUT xlsx_writer.book_r, p_id IN NUMBER, p_style number)
return xlsx_writer.book_r
is
  v_sheet number;

  l_datei_name        VARCHAR2(500);
  l_von               VARCHAR2(200);
  l_bis               VARCHAR2(200);
  l_getrimmter_mw     VARCHAR2(200);
  l_region            VARCHAR2(200);

  l_lvs_raw           VARCHAR2(4000);
  l_lvs_namen         VARCHAR2(4000);
  l_lieferant_raw     VARCHAR2(4000); 
  l_lieferant_namen   VARCHAR2(4000);

  l_vergabesumme      VARCHAR2(200); 

  l_names SYS.ODCIVARCHAR2LIST;
  l_vals  SYS.ODCIVARCHAR2LIST;

  font_db_bold    integer;
  font_white_bold integer;
  fill_black      integer;
  border_all      integer;
  style_bold      integer;
  style_header    integer;
  style_header_b  integer;
  style_data      integer;

  l_row_num     integer := 2; 

begin

  v_sheet  := XLSX_WRITER.add_sheet(p_ws, 'Eingabe');

  font_db_bold    := xlsx_writer.add_font(p_ws, 'DB Office', 10, b => true);
  font_white_bold := xlsx_writer.add_font(p_ws, 'DB Office', 10, color=> 'theme="0"', b => true);
  fill_black      := xlsx_writer.add_fill(p_ws, '<patternFill patternType="solid"><fgColor rgb="000000"/><bgColor indexed="64"/></patternFill>');
  border_all      := xlsx_writer.add_border(p_ws, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');

  style_bold     := xlsx_writer.add_cell_style(p_ws, font_id => font_db_bold, border_id => border_all);
  style_header   := xlsx_writer.add_cell_style(p_ws, font_id => font_white_bold, fill_id => fill_black, border_id => border_all);
  style_header_b := xlsx_writer.add_cell_style(p_ws, fill_id => fill_black, border_id => border_all);
  style_data     := xlsx_writer.add_cell_style(p_ws, border_id => border_all);

  xlsx_writer.col_width(p_ws, v_sheet, 1, 30);
  xlsx_writer.col_width(p_ws, v_sheet, 2, 100);

  BEGIN
      IF p_id IS NOT NULL THEN
          SELECT 
            NAME,
            TO_CHAR(VON, 'DD.MM.YYYY'),
            TO_CHAR(BIS, 'DD.MM.YYYY'),
            CASE WHEN GETRIMMTER_MITTELWERT IS NOT NULL THEN TO_CHAR(GETRIMMTER_MITTELWERT) || '%' ELSE NULL END,         
            TO_CHAR(REGION),
            TO_CHAR(LVS),
            NVL(LIEFERANT_LISTE, LIEFERANT),
            case 
              when VERGABESUMME is null and VERGABESUMME2 is null then null
              else TO_CHAR(VERGABESUMME) || ' bis ' || TO_CHAR(VERGABESUMME2)
            end
          INTO 
            l_datei_name, l_von, l_bis, l_getrimmter_mw, l_region, l_lvs_raw, l_lieferant_raw, l_vergabesumme
          FROM PD_IMPORT_X86
          WHERE id = p_id;

          IF l_lieferant_raw IS NOT NULL THEN
              BEGIN
                  SELECT LISTAGG(DISTINCT AUFTRAGNAHMER_NAME, ', ') WITHIN GROUP (ORDER BY AUFTRAGNAHMER_NAME)
                  INTO l_lieferant_namen
                  FROM PD_AUFTRAEGE
                  WHERE KEDITOREN_NUMMER IN (SELECT column_value FROM table(apex_string.split(l_lieferant_raw, ':')));
              EXCEPTION WHEN OTHERS THEN l_lieferant_namen := substr(l_lieferant_raw, 1, 4000);
              END;
          END IF;

          IF l_lvs_raw IS NOT NULL THEN
              BEGIN
                  SELECT LISTAGG(DISTINCT NAME, ', ') WITHIN GROUP (ORDER BY NAME)
                  INTO l_lvs_namen
                  FROM PD_MUSTER_LVS
                  WHERE ID IN (SELECT to_number(column_value) FROM table(apex_string.split(l_lvs_raw, ':')));
              EXCEPTION WHEN OTHERS THEN l_lvs_namen := substr(l_lvs_raw, 1, 4000);
              END;
          END IF;

      END IF;
  EXCEPTION
      WHEN NO_DATA_FOUND THEN NULL; 
  END;

  l_names := SYS.ODCIVARCHAR2LIST('Dateiname', 'Von', 'Bis', 'Getrimmter Mittelwert', 'Region', 'LVS', 'Lieferanten', 'Vergabesumme');
  l_vals := SYS.ODCIVARCHAR2LIST(l_datei_name, l_von, l_bis, l_getrimmter_mw, l_region, l_lvs_namen, l_lieferant_namen, l_vergabesumme);

  xlsx_writer.add_cell(p_ws, v_sheet, 1, 1, style_id => style_header, text => 'Eingabe-Parameter lauten:');
  xlsx_writer.add_cell(p_ws, v_sheet, 1, 2, style_id => style_header_b, text => '');

  FOR i IN 1 .. l_names.COUNT LOOP
    IF l_vals(i) IS NOT NULL THEN
        xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => l_names(i) || ':');
        xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => l_vals(i));
        l_row_num := l_row_num + 1;
    END IF;
  END LOOP;

  return p_ws;
end print_param_worksheet;
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function print_param_worksheet(
    p_ws IN OUT xlsx_writer.book_r, 
    p_von IN DATE, 
    p_bis IN DATE, 
    p_regionen IN VARCHAR2, 
    p_lvs_raw IN VARCHAR2, 
    p_liferant_raw IN VARCHAR2,
    p_trimm IN NUMBER DEFAULT NULL,
    p_vergabesumme_von IN NUMBER DEFAULT NULL, 
    p_vergabesumme_bis IN NUMBER DEFAULT NULL  
) return xlsx_writer.book_r 
is
  v_sheet number;
  l_lvs_namen         VARCHAR2(4000);
  l_lieferant_namen   VARCHAR2(4000);
  l_region_namen      VARCHAR2(4000);
  l_vergabesumme      VARCHAR2(400);

  font_db_bold    integer;
  font_white_bold integer;
  fill_black      integer;
  border_all      integer;
  style_bold      integer;
  style_header    integer;
  style_header_b  integer;
  style_data      integer;

  l_row_num     integer := 2; 

begin
  v_sheet  := XLSX_WRITER.add_sheet(p_ws, 'Eingabe');

  font_db_bold    := xlsx_writer.add_font(p_ws, 'DB Office', 10, b => true);
  font_white_bold := xlsx_writer.add_font(p_ws, 'DB Office', 10, color=> 'theme="0"', b => true);
  fill_black      := xlsx_writer.add_fill(p_ws, '<patternFill patternType="solid"><fgColor rgb="000000"/><bgColor indexed="64"/></patternFill>');
  border_all      := xlsx_writer.add_border(p_ws, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');

  style_bold     := xlsx_writer.add_cell_style(p_ws, font_id => font_db_bold, border_id => border_all);
  style_header   := xlsx_writer.add_cell_style(p_ws, font_id => font_white_bold, fill_id => fill_black, border_id => border_all);
  style_header_b := xlsx_writer.add_cell_style(p_ws, fill_id => fill_black, border_id => border_all);
  style_data     := xlsx_writer.add_cell_style(p_ws, border_id => border_all);

  xlsx_writer.col_width(p_ws, v_sheet, 1, 30);
  xlsx_writer.col_width(p_ws, v_sheet, 2, 100);

  IF p_liferant_raw IS NOT NULL THEN
      BEGIN
          SELECT LISTAGG(DISTINCT AUFTRAGNAHMER_NAME, ', ') WITHIN GROUP (ORDER BY AUFTRAGNAHMER_NAME)
          INTO l_lieferant_namen FROM PD_AUFTRAEGE
          WHERE KEDITOREN_NUMMER IN (SELECT column_value FROM table(apex_string.split(p_liferant_raw, ':')));
      EXCEPTION WHEN OTHERS THEN l_lieferant_namen := substr(p_liferant_raw, 1, 4000);
      END;
  END IF;

  IF p_lvs_raw IS NOT NULL THEN
      BEGIN
          SELECT LISTAGG(DISTINCT NAME, ', ') WITHIN GROUP (ORDER BY NAME)
          INTO l_lvs_namen FROM PD_MUSTER_LVS
          WHERE ID IN (SELECT to_number(column_value) FROM table(apex_string.split(p_lvs_raw, ':')));
      EXCEPTION WHEN OTHERS THEN l_lvs_namen := substr(p_lvs_raw, 1, 4000);
      END;
  END IF;

  IF p_regionen IS NOT NULL THEN
      BEGIN
          SELECT LISTAGG(DISTINCT NAME, ', ') WITHIN GROUP (ORDER BY NAME)
          INTO l_region_namen FROM PD_REGION
          WHERE ID IN (SELECT to_number(column_value) FROM table(apex_string.split(p_regionen, ':')));
      EXCEPTION WHEN OTHERS THEN l_region_namen := substr(p_regionen, 1, 4000);
      END;
  END IF;

  IF p_vergabesumme_von IS NOT NULL OR p_vergabesumme_bis IS NOT NULL THEN
      l_vergabesumme := TO_CHAR(p_vergabesumme_von) || ' bis ' || TO_CHAR(p_vergabesumme_bis);
  END IF;

  xlsx_writer.add_cell(p_ws, v_sheet, 1, 1, style_id => style_header, text => 'Eingabe-Parameter lauten:');
  xlsx_writer.add_cell(p_ws, v_sheet, 1, 2, style_id => style_header_b, text => '');

  IF p_von IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'Von:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => TO_CHAR(p_von, 'DD.MM.YYYY'));
      l_row_num := l_row_num + 1;
  END IF;

  IF p_bis IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'Bis:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => TO_CHAR(p_bis, 'DD.MM.YYYY'));
      l_row_num := l_row_num + 1;
  END IF;

  IF l_region_namen IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'Region:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => l_region_namen);
      l_row_num := l_row_num + 1;
  END IF;

  IF l_lvs_namen IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'LVS:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => l_lvs_namen);
      l_row_num := l_row_num + 1;
  END IF;

  IF l_lieferant_namen IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'Lieferanten:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => l_lieferant_namen);
      l_row_num := l_row_num + 1;
  END IF;

  IF l_vergabesumme IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'Vergabesumme:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => l_vergabesumme);
      l_row_num := l_row_num + 1;
  END IF;

  IF p_trimm IS NOT NULL THEN
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 1, style_id => style_bold, text => 'Getrimmter Mittelwert:');
      xlsx_writer.add_cell(p_ws, v_sheet, l_row_num, 2, style_id => style_data, text => p_trimm || ' %');
      l_row_num := l_row_num + 1;
  END IF;

  return p_ws;
end print_param_worksheet;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
FUNCTION GetNamespace(p_blob_id in number) return number as
  v_return number;
begin

  select  case
           /* when upper (name) like '%.X82%' 
                then 4

                */  
            when dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'), 1, 1) > 0 
              or dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/200407'), 1, 1) > 0  --x82 Format 28.02.2026
              or dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/DA82/'), 1, 1) > 0    --x82 Format 28.02.2026
              then 2
            -- CW 12.03.2025: Neuer Typ 3
             when dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/DA83/3.3'), 1, 1) > 0 
             --neuer Typ akram Sioud 13.02.2026
             or dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/DA86/3.3'), 1, 1) > 0  
             then 1

            else 3
          end
  into    v_return
  from    pd_import_x86
  where   id = p_blob_id;

  --  DBS_LOGGING.LOG_DEBUG_AT(v_return,'import');

  return v_return;

exception
  when others
    then return 1;
end;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE SendMailAuswertung(p_user_id number,p_anhang blob,p_filename varchar2) as
    v_mail                  dbs_email.rt_mail;
    r_mail_attachment       dbs_email.rt_attachment;
    r_mail_attachment_null  dbs_email.rt_attachment;
    t_mail_attachments      dbs_email.tt_attachments := dbs_email.tt_attachments();
    v_attachment_cnt        pls_integer := 0;
    v_mail_id               number;
    v_workspace_id          number;
begin

    v_workspace_id := apex_util.find_security_group_id (p_workspace => 'PREISDB');
    apex_util.set_security_group_id (p_security_group_id => v_workspace_id);

    --Empfänger-Mailadresse holen
    select  mail 
    into    v_mail.send_to
    from    dbs_user 
    where   id = p_user_id; 

    --Mail versenden 
    v_mail_id := apex_mail.send(
        p_to        => v_mail.send_to,
        p_from      => 'noreply@deutschebahn.com',
        p_body      => 'Sehr geehrte(r) Anwender(in),<br><br>anbei erhalten Sie Ihre gewünschte Auswertung aus der Preisdatenbank.',
        p_body_html => 'Sehr geehrte(r) Anwender(in),<br><br>anbei erhalten Sie Ihre gewünschte Auswertung aus der Preisdatenbank.',
        p_subj      => 'Ihre Auswertung aus der Preisdatenbank',
        p_cc        => null,
        p_bcc       => null,
        p_replyto   => null
        );

    --Anhang der Mail hinzufügen
    r_mail_attachment.content := p_anhang;
    r_mail_attachment.file_name := p_filename;
    t_mail_attachments.extend;
    t_mail_attachments(t_mail_attachments.last) := r_mail_attachment;

    for i in 1..t_mail_attachments.count loop
        v_attachment_cnt := v_attachment_cnt + 1;
        r_mail_attachment := r_mail_attachment_null;
        r_mail_attachment := t_mail_attachments(i);

        if r_mail_attachment.filebrowse_value is not null then
            select blob_content, filename
            into r_mail_attachment.content, r_mail_attachment.file_name
            from apex_application_temp_files
            where name = r_mail_attachment.filebrowse_value;
        elsif r_mail_attachment.file_name is null then
                r_mail_attachment.file_name := 'ATT' || to_char(v_attachment_cnt, 'FM009');
        end if;

        apex_mail.add_attachment(
            p_mail_id       => v_mail_id,
            p_attachment    => r_mail_attachment.content,
            p_filename      => r_mail_attachment.file_name,
            p_mime_type     => 'application/octet-stream'
        );
    end loop;

    apex_mail.push_queue;

end SendMailAuswertung;

PROCEDURE SendMailAuswertungFehler(p_user_id number,p_error varchar2) as
    v_mail                  dbs_email.rt_mail;
    r_mail_attachment       dbs_email.rt_attachment;
    r_mail_attachment_null  dbs_email.rt_attachment;
    t_mail_attachments      dbs_email.tt_attachments := dbs_email.tt_attachments();
    v_attachment_cnt        pls_integer := 0;
    v_mail_id               number;
    v_workspace_id          number;
begin

    v_workspace_id := apex_util.find_security_group_id (p_workspace => 'PREISDB');
    apex_util.set_security_group_id (p_security_group_id => v_workspace_id);

    --Empfänger-Mailadresse holen
    select  mail 
    into    v_mail.send_to
    from    dbs_user 
    where   id = p_user_id; 

    --Mail versenden 
    v_mail_id := apex_mail.send(
        p_to        => v_mail.send_to,
        p_from      => 'noreply@deutschebahn.com',
        p_body      => p_error,
        p_body_html => p_error,
        p_subj      => 'Ihre Auswertung aus der Preisdatenbank (Fehler)',
        p_cc        => null,
        p_bcc       => null,
        p_replyto   => null
        );

    apex_mail.push_queue;

end SendMailAuswertungFehler;

END PREISDATENBANK_PKG;