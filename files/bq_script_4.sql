-- BLOCK 1
-- complexity_score: 0
-- used_llm: False
CREATE OR REPLACE EDITIONABLE PROCEDURE "VNADSPRD"."SP_IEN_ICON_TBLS_REFRESH" AS
    v_icoe_equipment_tbl   NUMBER

-- BLOCK 2
-- complexity_score: 0
-- used_llm: False
v_icoe_card_tbl        NUMBER

-- BLOCK 3
-- complexity_score: 0
-- used_llm: False
v_icoe_slot_tbl        NUMBER

-- BLOCK 4
-- complexity_score: 0
-- used_llm: False
v_icoe_site_tbl        NUMBER

-- BLOCK 5
-- complexity_score: 1
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
BEGIN

    dbms_output.put_line('Step 1 start : create ICOE_EQUIPMENT_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 6
-- complexity_score: 0
-- used_llm: False
-- note: Qualified simple table references.
SELECT
        COUNT(*)
    INTO `JSON_TO_BQ.BQ_DATASET.v_icoe_equipment_tbl`
    FROM
        `JSON_TO_BQ.BQ_DATASET.all_tables`
    WHERE
        table_name = 'ICOE_EQUIPMENT_TBL'

-- BLOCK 7
-- complexity_score: 4
-- used_llm: True
IF v_icoe_equipment_tbl = 1 THEN
  EXECUTE IMMEDIATE 'DROP TABLE IF EXISTS `VNADSPRD.ICOE_EQUIPMENT_TBL`';
END IF;

-- BLOCK 8
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE OR REPLACE TABLE `VNADSPRD.ICOE_EQUIPMENT_TBL` AS
SELECT *
FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD.icoe_mv_logical_equipment`;

-- BLOCK 9
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'GRANT SELECT ON `VNADSPRD.ICOE_EQUIPMENT_TBL` TO `VNA_SELECT`';

-- BLOCK 10
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE OR REPLACE TABLE `VNADSPRD.ICOE_EQUIPMENT_TBL` AS
SELECT *
FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD.icoe_mv_logical_equipment`;

-- BLOCK 11
-- complexity_score: 4
-- used_llm: True
GRANT SELECT ON TABLE VNADSPRD.ICOE_EQUIPMENT_TBL TO ROLE 'VNA_SELECT';

-- BLOCK 12
-- complexity_score: 0
-- used_llm: False
END IF

-- BLOCK 13
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'CREATE INDEX VNADSPRD.IDX_ICOE_EQUIP_TBL ON VNADSPRD.ICOE_EQUIPMENT_TBL(PHYSICAL_EQUIPMENT_REFERENC_ID) NOLOGGING'

-- BLOCK 14
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'CREATE INDEX VNADSPRD.IDX1_ICOE_EQUIP_TBL ON VNADSPRD.ICOE_EQUIPMENT_TBL(PARENT_EQP_REFERENCE_ID) NOLOGGING'

-- BLOCK 15
-- complexity_score: 4
-- used_llm: True
-- Index creation not supported in BigQuery; consider clustering on SHELF_TYPE if needed.

-- BLOCK 16
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'ALTER TABLE `VNADSPRD.ICOE_EQUIPMENT_TBL` SET OPTIONS (clustering_fields=[''SITE_REFERENCE_ID'']);';

-- BLOCK 17
-- complexity_score: 4
-- used_llm: True
-- Index creation is not supported in BigQuery. If clustering or partitioning is required, apply the appropriate table options instead.

-- BLOCK 18
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 1 end : create ICOE_EQUIPMENT_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 19
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 2 start : create ICOE_CARD_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 20
-- complexity_score: 0
-- used_llm: False
-- note: Qualified simple table references.
SELECT
        COUNT(*)
    INTO `JSON_TO_BQ.BQ_DATASET.v_icoe_card_tbl`
    FROM
        `JSON_TO_BQ.BQ_DATASET.all_tables`
    WHERE
        table_name = 'ICOE_CARD_TBL'

-- BLOCK 21
-- complexity_score: 4
-- used_llm: True
DECLARE v_icoe_card_tbl INT64;
SET v_icoe_card_tbl = 1;  -- set to 1 to trigger the drop, adjust as needed

IF v_icoe_card_tbl = 1 THEN
  EXECUTE IMMEDIATE 'DROP TABLE IF EXISTS `VNADSPRD.ICOE_CARD_TBL`';
END IF;

-- BLOCK 22
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE OR REPLACE TABLE VNADSPRD.ICOE_CARD_TBL AS
SELECT * FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD.icoe_mv_logical_card`;

-- BLOCK 23
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE "GRANT SELECT ON TABLE `VNADSPRD.ICOE_CARD_TBL` TO ROLE 'VNA_SELECT';"

-- BLOCK 24
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
EXECUTE IMMEDIATE 'CREATE TABLE VNADSPRD.ICOE_CARD_TBL AS SELECT * FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.icoe_mv_logical_card';

-- BLOCK 25
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'GRANT SELECT ON `VNADSPRD.ICOE_CARD_TBL` TO `VNA_SELECT`';

-- BLOCK 26
-- complexity_score: 0
-- used_llm: False
END IF

-- BLOCK 27
-- complexity_score: 4
-- used_llm: True
-- Index creation is not supported in BigQuery. If clustering on CARD_REFERENCE_ID is desired, consider adding clustering fields to the table definition.

-- BLOCK 28
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'CREATE INDEX VNADSPRD.IDX1_ICOE_CARD_TBL ON VNADSPRD.ICOE_CARD_TBL(SLOT_REFERENCE_ID) NOLOGGING'

-- BLOCK 29
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'CREATE INDEX VNADSPRD.IDX2_ICOE_CARD_TBL ON VNADSPRD.ICOE_CARD_TBL(PARENT_CARD_REF_ID) NOLOGGING'

-- BLOCK 30
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 2 end : create ICOE_CARD_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 31
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 3 start : create ICOE_SLOT_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 32
-- complexity_score: 0
-- used_llm: False
-- note: Qualified simple table references.
SELECT
        COUNT(*)
    INTO `JSON_TO_BQ.BQ_DATASET.v_icoe_slot_tbl`
    FROM
        `JSON_TO_BQ.BQ_DATASET.all_tables`
    WHERE
        table_name = 'ICOE_SLOT_TBL'

-- BLOCK 33
-- complexity_score: 4
-- used_llm: True
BEGIN
  IF v_icoe_slot_tbl = 1 THEN
    EXECUTE IMMEDIATE 'DROP TABLE `VNADSPRD.ICOE_SLOT_TBL`';
  END IF;
END;

-- BLOCK 34
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE OR REPLACE TABLE VNADSPRD.ICOE_SLOT_TBL AS
SELECT *
FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.ICOE_PVNR_MV_LOGICAL_SLOT;

-- BLOCK 35
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'GRANT SELECT ON `VNADSPRD.ICOE_SLOT_TBL` TO `VNA_SELECT`';

-- BLOCK 36
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE OR REPLACE TABLE VNADSPRD.ICOE_SLOT_TBL AS
SELECT * FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.ICOE_PVNR_MV_LOGICAL_SLOT;

-- BLOCK 37
-- complexity_score: 4
-- used_llm: True
GRANT SELECT ON `VNADSPRD.ICOE_SLOT_TBL` TO `VNA_SELECT`;

-- BLOCK 38
-- complexity_score: 0
-- used_llm: False
END IF

-- BLOCK 39
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'CREATE INDEX VNADSPRD.IDX_SLOT_TBL ON VNADSPRD.ICOE_SLOT_TBL(SLOT_REFERENCE_ID) NOLOGGING'

-- BLOCK 40
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 3 end : create ICOE_SLOT_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 41
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 4 start : create ICOE_SITE_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 42
-- complexity_score: 0
-- used_llm: False
-- note: Qualified simple table references.
SELECT
        COUNT(*)
    INTO `JSON_TO_BQ.BQ_DATASET.v_icoe_site_tbl`
    FROM
        `JSON_TO_BQ.BQ_DATASET.all_tables`
    WHERE
        table_name = 'ICOE_SITE_TBL'

-- BLOCK 43
-- complexity_score: 4
-- used_llm: True
BEGIN
  IF v_icoe_site_tbl = 1 THEN
    EXECUTE IMMEDIATE 'DROP TABLE IF EXISTS `VNADSPRD.ICOE_SITE_TBL`';
  END IF;
END;

-- BLOCK 44
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE OR REPLACE TABLE VNADSPRD.ICOE_SITE_TBL AS
SELECT *
FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD.ICOE_MV_SITE`;

-- BLOCK 45
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'GRANT SELECT ON TABLE `VNADSPRD`.`ICOE_SITE_TBL` TO ROLE ''VNA_SELECT''';

-- BLOCK 46
-- complexity_score: 4
-- used_llm: True
-- note: Qualified simple table references.
CREATE TABLE VNADSPRD.ICOE_SITE_TBL AS SELECT * FROM `JSON_TO_BQ.BQ_DATASET.VNADSPRD.ICOE_MV_SITE`;

-- BLOCK 47
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'GRANT SELECT ON TABLE `VNADSPRD.ICOE_SITE_TBL` TO ROLE "VNA_SELECT"';

-- BLOCK 48
-- complexity_score: 0
-- used_llm: False
END IF

-- BLOCK 49
-- complexity_score: 4
-- used_llm: True
-- BigQuery does not support CREATE INDEX; this statement is retained as a comment for reference.
-- CREATE INDEX VNADSPRD.IDX_SITE_TBL ON VNADSPRD.ICOE_SITE_TBL(SITE_REFERENCE_ID) NOLOGGING;

-- BLOCK 50
-- complexity_score: 4
-- used_llm: True
-- BigQuery does not support CREATE INDEX. Equivalent performance tuning can be achieved via clustering or partitioning on the target table.
-- The following statement is a no-op placeholder to maintain script structure.
-- CREATE INDEX VNADSPRD.IDX1_SITE_TBL ON VNADSPRD.ICOE_SITE_TBL(CLLI) NOLOGGING;

-- BLOCK 51
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 4 end  : create ICOE_SITE_TBL temp table   - ' || CURRENT_TIMESTAMP())

-- BLOCK 52
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 5 start  : gather stats on temp tables   - ' || CURRENT_TIMESTAMP())

-- BLOCK 53
-- complexity_score: 0
-- used_llm: False
dbms_stats.gather_table_stats(ownname => 'VNADSPRD', tabname => 'ICOE_CARD_TBL', degree => 16, method_opt => 'FOR ALL COLUMNS SIZE 1'
    , cascade => true)

-- BLOCK 54
-- complexity_score: 0
-- used_llm: False
dbms_stats.gather_table_stats(ownname => 'VNADSPRD', tabname => 'ICOE_SLOT_TBL', degree => 16, method_opt => 'FOR ALL COLUMNS SIZE 1'
    , cascade => true)

-- BLOCK 55
-- complexity_score: 0
-- used_llm: False
dbms_stats.gather_table_stats(ownname => 'VNADSPRD', tabname => 'ICOE_SITE_TBL', degree => 16, method_opt => 'FOR ALL COLUMNS SIZE 1'
    , cascade => true)

-- BLOCK 56
-- complexity_score: 0
-- used_llm: False
dbms_stats.gather_table_stats(ownname => 'VNADSPRD', tabname => 'ICOE_EQUIPMENT_TBL', degree => 16, method_opt => 'FOR ALL COLUMNS SIZE 1'
    , cascade => true)

-- BLOCK 57
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
dbms_output.put_line('Step 5 end  : gather stats on temp tables   - ' || CURRENT_TIMESTAMP())

-- BLOCK 58
-- complexity_score: 5
-- used_llm: True
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
-- note: Detected Oracle outer join markers and annotated them for LLM conversion.
-- note: Qualified simple table references.
dbms_output.put_line('

	   EXECUTE IMMEDIATE 'TRUNCATE TABLE vnadsprd.IEN_MR_CHANNEL_TEMP';
        EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX1_IEN_MR_CHANNEL_TEMP UNUSABLE';
        EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX2_IEN_MR_CHANNEL_TEMP UNUSABLE';

      INSERT  INTO `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ien_mr_channel_temp (
            trail_id,
            trail_name,
            type,
            status,
            version,
            parent_trail_id,
            channel_name,
            project_id,
            aclli,
            zclli

        )
		SELECT VW.TRAIL_ID,

               VW.TRAIL_NAME,

               VW.TYPE,

               VW.STATUS,

               VW.VERSION,

               PARENT_TRAIL_ID,

               MIN((CHANNEL_NAME)) CHANNEL_NAME,

               VW.PROJECT_ID,

               ASITE.CLLI ACLLI,

               ZSITE.CLLI ZCLLI

          FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL_ELEMENT TE, vnadsprd.NAUT_TRAIL VW,

          vnadsprd.ICOE_SITE_TBL ASITE,

           vnadsprd.ICOE_SITE_TBL ZSITE

         WHERE

            VW.TRAIL_ID = TE.TRAIL_ID

           AND TE.ELEMENT_TYPE = 'P'

           AND VW.A_SITE_ID =ASITE.SITE_REFERENCE_ID /*ORACLE_OUTER_JOIN_MARKER*/

           AND VW.Z_SITE_ID =ZSITE.SITE_REFERENCE_ID /*ORACLE_OUTER_JOIN_MARKER*/

         GROUP BY VW.TRAIL_ID,

                  VW.TRAIL_NAME,

                  VW.TYPE,

                  VW.STATUS,

                  VW.VERSION,

                  TE.PARENT_TRAIL_ID,

                  VW.PROJECT_ID,

                  ASITE.CLLI,

                  ZSITE.CLLI;
        COMMIT;
        EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX1_IEN_MR_CHANNEL_TEMP REBUILD';
        EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX2_IEN_MR_CHANNEL_TEMP REBUILD';
   dbms_output.put_line('

  dbms_output.put_line('Step 6 start  : truncate and insert IEN_MARKET_INPUT_LIST_STG   - ' || CURRENT_TIMESTAMP());

        EXECUTE IMMEDIATE 'truncate table vnadsprd.IEN_MARKET_INPUT_LIST_STG';
		 EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX1_IEN_MARKET_INPUT_LIST_STG UNUSABLE';
		 EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX2_IEN_MARKET_INPUT_LIST_STG UNUSABLE';
		EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX3_IEN_MARKET_INPUT_LIST_STG UNUSABLE';
        INSERT INTO `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_MARKET_INPUT_LIST_STG (TRAIL_ID	,
	TRAIL_NAME	,
	STATUS	,
	TYPE	,
	VERSION	,
	SEQUENCE	,
	PORT_REFERENCE_ID	,
	SOURCE	,
	CHANNEL_NAME	,
	SEQUENCE_NUMBER	,
	PORT_STATUS	,
	ELEMENT_TYPE,
	NF_ID,
	ACLLI,
	ZCLLI,
	LAST_REFRESHED_TS,
	EQUIPMENT_ID )
SELECT DISTINCT VW.TRAIL_ID,
                VW.TRAIL_NAME,
                VW.STATUS,
                VW.TYPE,
                VW.VERSION,
                TCE.SEQUENCE,
                E.ELEMENT_REF_ID AS PORT_REFERENCE_ID,
                E.SOURCE,
                E.CHANNEL_NAME,
                NULL AS SEQUENCE_NUMBER,
                NULL AS PORT_STATUS,
                E.ELEMENT_TYPE,
                VW.PROJECT_ID,
                ASITE.CLLI ACLLI,
                ZSITE.CLLI ZCLLI,
                CURRENT_TIMESTAMP(),
				NULL AS EQUIPMENT_ID
  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL_ELEMENT E,
       vnadsprd.NAUT_TRAIL   VW,
       vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,

       vnadsprd.ICOE_SITE_TBL ASITE,

           vnadsprd.ICOE_SITE_TBL ZSITE
   WHERE VW.TRAIL_ID = E.TRAIL_ID
   AND E.ELEMENT_TYPE = 'E'
   AND E.SOURCE IN ('IVAPP_PORT', 'IVAPP_LOGICAL', 'IVAPP_PANEL')
   AND TCE.ELEMENT_ID=E.ELEMENT_ID
   AND VW.A_SITE_ID =ASITE.SITE_REFERENCE_ID /*ORACLE_OUTER_JOIN_MARKER*/
   AND VW.Z_SITE_ID =ZSITE.SITE_REFERENCE_ID /*ORACLE_OUTER_JOIN_MARKER*/
UNION
SELECT X.TRAIL_ID,
       X.TRAIL_NAME,
       X.STATUS,
       X.TYPE,
       X.VERSION,
       TCE.SEQUENCE,
       T.A_PORT_ID PORT_REFERENCE_ID,
       'NAUTILUS' AS SOURCE,
       TC.CHANNEL_NAME,
       TC.SEQUENCE_NUMBER,
       TC.STATUS PORT_STATUS,
       TE.ELEMENT_TYPE,
       X.PROJECT_ID,
       X.ACLLI,
       X.ZCLLI,
       CURRENT_TIMESTAMP(),
	   NULL AS EQUIPMENT_ID
  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,
       vnadsprd.NAUT_TRAIL_CHANNEL TC,
       vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,
       vnadsprd.NAUT_TRAIL_ELEMENT TE,
	   vnadsprd.ien_mr_channel_temp X

 WHERE T.TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.PARENT_TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.CHANNEL_NAME = to_char(X.CHANNEL_NAME)
   AND TE.CHANNEL_NAME= to_char(X.CHANNEL_NAME)
   AND TE.PARENT_TRAIL_ID= X.PARENT_TRAIL_ID
   AND TCE.ELEMENT_ID=TE.ELEMENT_ID
UNION
SELECT X.TRAIL_ID,
       X.TRAIL_NAME,
       X.STATUS,
       X.TYPE,
       X.VERSION,
       TCE.SEQUENCE,
       T.Z_PORT_ID PORT_REFERENCE_ID,
       'NAUTILUS' AS SOURCE,
       TC.CHANNEL_NAME,
       TC.SEQUENCE_NUMBER,
       TC.STATUS PORT_STATUS,
       TE.ELEMENT_TYPE,
       X.PROJECT_ID,
       X.ACLLI,
       X.ZCLLI,
       CURRENT_TIMESTAMP(),
	   NULL AS EQUIPMENT_ID
  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,
       vnadsprd.NAUT_TRAIL_CHANNEL TC,
        vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,
       vnadsprd.NAUT_TRAIL_ELEMENT TE,
	   vnadsprd.ien_mr_channel_temp X

 WHERE T.TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.PARENT_TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.CHANNEL_NAME = to_char(X.CHANNEL_NAME)
    AND TE.CHANNEL_NAME= to_char(X.CHANNEL_NAME)
   AND TE.PARENT_TRAIL_ID= X.PARENT_TRAIL_ID
   AND TCE.ELEMENT_ID=TE.ELEMENT_ID
UNION
   SELECT X.TRAIL_ID,
       X.TRAIL_NAME,
       X.STATUS,
       X.TYPE,
       X.VERSION,
       TCE.SEQUENCE,
       T.A_PORT_ID_REVERSE PORT_REFERENCE_ID,
       'NAUTILUS' AS SOURCE,
       TC.CHANNEL_NAME,
       TC.SEQUENCE_NUMBER,
       TC.STATUS PORT_STATUS,
       TE.ELEMENT_TYPE,
       X.PROJECT_ID,
       X.ACLLI,
       X.ZCLLI,
       CURRENT_TIMESTAMP(),
	   NULL AS EQUIPMENT_ID
  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,
       vnadsprd.NAUT_TRAIL_CHANNEL TC,
        vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,
       vnadsprd.NAUT_TRAIL_ELEMENT TE,
	   vnadsprd.ien_mr_channel_temp X

 WHERE T.TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.PARENT_TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.CHANNEL_NAME = to_char(X.CHANNEL_NAME)
    AND TE.CHANNEL_NAME = to_char(X.CHANNEL_NAME)
   AND TE.PARENT_TRAIL_ID= X.PARENT_TRAIL_ID
   AND TCE.ELEMENT_ID=TE.ELEMENT_ID
   UNION
   SELECT X.TRAIL_ID,
       X.TRAIL_NAME,
       X.STATUS,
       X.TYPE,
       X.VERSION,
       TCE.SEQUENCE,
       T.Z_PORT_ID_REVERSE PORT_REFERENCE_ID,
       'NAUTILUS' AS SOURCE,
       TC.CHANNEL_NAME,
       TC.SEQUENCE_NUMBER,
       TC.STATUS PORT_STATUS,
       TE.ELEMENT_TYPE,
       X.PROJECT_ID,
       X.ACLLI,
       X.ZCLLI,
       CURRENT_TIMESTAMP(),
	   NULL AS EQUIPMENT_ID
  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,
       vnadsprd.NAUT_TRAIL_CHANNEL TC,
        vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,
       vnadsprd.NAUT_TRAIL_ELEMENT TE,
	    vnadsprd.ien_mr_channel_temp X

 WHERE T.TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.PARENT_TRAIL_ID = X.PARENT_TRAIL_ID
   AND TC.CHANNEL_NAME = to_char(X.CHANNEL_NAME)
    AND TE.CHANNEL_NAME= to_char(X.CHANNEL_NAME)
   AND TE.PARENT_TRAIL_ID= X.PARENT_TRAIL_ID
   AND TCE.ELEMENT_ID=TE.ELEMENT_ID;
   COMMIT;
 EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX1_IEN_MARKET_INPUT_LIST_STG REBUILD';
 EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX2_IEN_MARKET_INPUT_LIST_STG REBUILD';
 EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IDX3_IEN_MARKET_INPUT_LIST_STG REBUILD';
  dbms_output.put_line('Step 6 end  : truncate and insert IEN_MARKET_INPUT_LIST_STG   - ' || CURRENT_TIMESTAMP());

  dbms_output.put_line('Step 7 start  : truncate and insert IEN_MARKET_INPUT_LIST_STG_TMP   - ' || CURRENT_TIMESTAMP());
  EXECUTE IMMEDIATE 'truncate table vnadsprd.IEN_MARKET_INPUT_LIST_STG_TMP';
    INSERT INTO `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_MARKET_INPUT_LIST_STG_TMP (TRAIL_ID,
	TRAIL_NAME	,
	STATUS	,
	TYPE	,
	VERSION	,
	SEQUENCE,
	PORT_REFERENCE_ID,
	SOURCE	,
	CHANNEL_NAME,
	SEQUENCE_NUMBER	,
	PORT_STATUS	,
	ELEMENT_TYPE,
	NF_ID,
	ACLLI,
	ZCLLI,
	LAST_REFRESHED_TS,
	EQUIPMENT_ID )
            SELECT DISTINCT
                VW.TRAIL_ID,
                VW.TRAIL_NAME,
                VW.STATUS,
                VW.TYPE,
                VW.VERSION,
                TCE.SEQUENCE,
                NULL          PORT_REFERENCE_ID,
                E.SOURCE,
                VW.A_PORT_AID,
                NULL          AS SEQUENCE_NUMBER,
                VW.STATUS     AS PORT_STATUS,
                E.ELEMENT_TYPE,
                VW.PROJECT_ID,
                ASITE.CLLI    ACLLI,
                ZSITE.CLLI    ZCLLI,
                CURRENT_TIMESTAMP(),
			    VW.A_EQUIPMENT_ID
            FROM
                `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.NAUT_TRAIL_ELEMENT           E,
                VNADSPRD.NAUT_TRAIL_TMP               VW,
                VNADSPRD.NAUT_TRAIL_COMPONENT_ELEMENT TCE,
                VNADSPRD.ICOE_SITE_TBL                ASITE,
                VNADSPRD.ICOE_SITE_TBL                ZSITE
            WHERE
                    VW.TRAIL_ID = E.TRAIL_ID
                AND E.ELEMENT_TYPE = 'K'
                AND UPPER(VW.TYPE) LIKE '%LAG%'
                AND TCE.ELEMENT_ID = E.ELEMENT_ID
                AND VW.A_SITE_ID = ASITE.SITE_REFERENCE_ID  /*ORACLE_OUTER_JOIN_MARKER*/
                AND VW.Z_SITE_ID = ZSITE.SITE_REFERENCE_ID  /*ORACLE_OUTER_JOIN_MARKER*/
            UNION
            SELECT DISTINCT
                VW.TRAIL_ID,
                VW.TRAIL_NAME,
                VW.STATUS,
                VW.TYPE,
                VW.VERSION,
                TCE.SEQUENCE,
                NULL          PORT_REFERENCE_ID,
                E.SOURCE,
                VW.Z_PORT_AID,
                NULL          AS SEQUENCE_NUMBER,
                VW.STATUS     AS PORT_STATUS,
                E.ELEMENT_TYPE,
                VW.PROJECT_ID,
                ASITE.CLLI    ACLLI,
                ZSITE.CLLI    ZCLLI,
                CURRENT_TIMESTAMP(),
				VW.Z_EQUIPMENT_ID
            FROM
                `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.NAUT_TRAIL_ELEMENT           E,
                VNADSPRD.NAUT_TRAIL_TMP               VW,
                VNADSPRD.NAUT_TRAIL_COMPONENT_ELEMENT TCE,
                VNADSPRD.ICOE_SITE_TBL                ASITE,
                VNADSPRD.ICOE_SITE_TBL                ZSITE
            WHERE
                    VW.TRAIL_ID = E.TRAIL_ID
                AND E.ELEMENT_TYPE = 'K'
                AND UPPER(VW.TYPE) LIKE '%LAG%'
                AND TCE.ELEMENT_ID = E.ELEMENT_ID
                AND VW.A_SITE_ID = ASITE.SITE_REFERENCE_ID  /*ORACLE_OUTER_JOIN_MARKER*/
                AND VW.Z_SITE_ID = ZSITE.SITE_REFERENCE_ID  /*ORACLE_OUTER_JOIN_MARKER*/ ;
        COMMIT;
  dbms_output.put_line('Step 7 end  : truncate and insert IEN_MARKET_INPUT_LIST_STG_TMP   - ' || CURRENT_TIMESTAMP());

  dbms_output.put_line('Step 8 start  : truncate and insert IEN_MARKET_INPUT_LIST   - ' || CURRENT_TIMESTAMP());

        EXECUTE IMMEDIATE 'TRUNCATE TABLE vnadsprd.IEN_MARKET_INPUT_LIST';
        EXECUTE IMMEDIATE 'ALTER INDEX vnadsprd.IEN_TRAIL_NAME_IDX UNUSABLE';
		dbms_output.put_line('vnadsprd.IEN_MARKET_INPUT_LIST tbl load started ' || to_char(CURRENT_TIMESTAMP(), 'DD-MON-YYYY HH24:MI:SS'));
        INSERT INTO `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.IEN_MARKET_INPUT_LIST (
TRAIL_ID,
MARKET_NAME	,
MARKET_STATE	,
CLLI	,
TRAIL_NAME	,
TID_LOGICAL	,
ALTERNATE_NAME	,
NFID	,
CREATED_TS
)
select distinct
TRAIL_ID,
IFNULL(market.VZB_MARKET_CITY_OR_VZT_REGION,'Not Available') market_name,
case when market.STATE is not null then market.STATE
else IFNULL(SUBSTR(S.CLLI,5,2),'N/A') end market_area,
IFNULL(S.CLLI,'N/A') AS EQUIP_CLLI,
G.TRAIL_NAME AS TRAIl_NAME,
(case when (E.SHELF_TYPE = 'MSERI') then MSE.TID_LOGICAL else E.TID_LOGICAL end) AS TID_LOGICAL,
(case when (E.SHELF_TYPE = 'MSERI') then MSE.ALTERNATE_NAME else E.ALTERNATE_NAME end) AS ALTERNATE_NAME,
IFNULL(G.NF_ID,'N/A') NF_ID,
CURRENT_TIMESTAMP()
from (((((select * from `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_MARKET_INPUT_LIST_STG where  element_type <> 'K') G join
`JSON_TO_BQ.BQ_DATASET.vnadsprd`.icoe_pvnr_t_logical_port P on((P.PORT_REFERENCE_ID = G.PORT_REFERENCE_ID)))

join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_EQUIPMENT_TBL E on((P.EQP_REFERENCE_ID = E.EQP_REFERENCE_ID)))

left join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_EQUIPMENT_TBL MSE on((E.PHYSICAL_EQUIPMENT_REFERENC_ID = MSE.EQP_REFERENCE_ID)))

join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_SITE_TBL S on((E.SITE_REFERENCE_ID = S.SITE_REFERENCE_ID)))
left join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_UT_PMO_TRACKER MARKET on (MARKET.SITE_CLLI = S.CLLI);
COMMIT;
  dbms_output.put_line('Step 8 end  : truncate and insert1 IEN_MARKET_INPUT_LIST   - ' || CURRENT_TIMESTAMP());

    dbms_output.put_line('Step 9 start  : gather stats on  IEN_MARKET_INPUT_LIST   - ' || CURRENT_TIMESTAMP());

BEGIN
                dbms_stats.gather_table_stats(ownname => 'VNADSPRD', tabname => 'IEN_MARKET_INPUT_LIST', degree => 16, method_opt
                => 'FOR ALL COLUMNS SIZE 1', cascade => true);
            END;

    dbms_output.put_line('Step 9 end  : gather stats on  IEN_MARKET_INPUT_LIST   - ' || CURRENT_TIMESTAMP());

  dbms_output.put_line('Step 10 start  : 2nd insert on IEN_MARKET_INPUT_LIST   - ' || CURRENT_TIMESTAMP());
INSERT INTO `JSON_TO_BQ.BQ_DATASET.VNADSPRD`.IEN_MARKET_INPUT_LIST (
TRAIL_ID,
MARKET_NAME	,
MARKET_STATE	,
CLLI	,
TRAIL_NAME	,
TID_LOGICAL	,
ALTERNATE_NAME	,
NFID	,
CREATED_TS
)
select  DISTINCT
TRAIL_ID,
IFNULL(market.VZB_MARKET_CITY_OR_VZT_REGION,'Not Available') market_name,
case when market.STATE is not null then market.STATE
else IFNULL(SUBSTR(S.CLLI,5,2),'N/A') end market_area,
IFNULL(S.CLLI,'N/A') AS EQUIP_CLLI,
G.TRAIL_NAME AS TRAIl_NAME,
(case when (E.SHELF_TYPE = 'MSERI') then MSE.TID_LOGICAL else E.TID_LOGICAL end) AS TID_LOGICAL,
(case when (E.SHELF_TYPE = 'MSERI') then MSE.ALTERNATE_NAME else E.ALTERNATE_NAME end) AS ALTERNATE_NAME,
IFNULL(G.NF_ID,'N/A') NF_ID,
CURRENT_TIMESTAMP()
from ((((select * from `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_MARKET_INPUT_LIST_STG_TMP where  element_type = 'K') G
join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_EQUIPMENT_TBL E on((E.EQP_REFERENCE_ID = G.EQUIPMENT_ID)))
left join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_EQUIPMENT_TBL MSE on((E.PHYSICAL_EQUIPMENT_REFERENC_ID = MSE.EQP_REFERENCE_ID)))
join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_SITE_TBL S on((E.SITE_REFERENCE_ID = S.SITE_REFERENCE_ID)))
left join `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_UT_PMO_TRACKER MARKET on (MARKET.SITE_CLLI = S.CLLI);
 COMMIT;
  dbms_output.put_line('Step 10 end  : 2nd insert on IEN_MARKET_INPUT_LIST   - ' || CURRENT_TIMESTAMP());

	EXECUTE IMMEDIATE 'ALTER INDEX VNADSPRD.IEN_TRAIL_NAME_IDX REBUILD';
END;
