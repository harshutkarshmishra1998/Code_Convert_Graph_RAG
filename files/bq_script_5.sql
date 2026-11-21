-- BLOCK 1
-- complexity_score: 5
-- used_llm: True
BEGIN
  TRUNCATE TABLE vnadsprd.IEN_MARKET_INPUT_LIST_STG;
END;

-- BLOCK 2
-- complexity_score: 4
-- used_llm: True
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
-- note: Detected Oracle outer join markers and annotated them for LLM conversion.
-- note: Qualified simple table references.
EXECUTE IMMEDIATE 'INSERT INTO `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_MARKET_INPUT_LIST_STG (TRAIL_ID	,

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

	LAST_REFRESHED_TS )

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

                CURRENT_TIMESTAMP()

  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL_ELEMENT E,

       vnadsprd.NAUT_TRAIL   VW,

       vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,

       vnadsprd.ICOE_SITE_TBL ASITE,

           vnadsprd.ICOE_SITE_TBL ZSITE

   WHERE VW.TRAIL_ID = E.TRAIL_ID

   AND E.ELEMENT_TYPE = ''E''

   AND E.SOURCE IN (''IVAPP_PORT'', ''IVAPP_LOGICAL'', ''IVAPP_PANEL'')

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

       ''NAUTILUS'' AS SOURCE,

       TC.CHANNEL_NAME,

       TC.SEQUENCE_NUMBER,

       TC.STATUS PORT_STATUS,

       TE.ELEMENT_TYPE,

       X.PROJECT_ID,

       X.ACLLI,

       X.ZCLLI,

       CURRENT_TIMESTAMP()

  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,

       vnadsprd.NAUT_TRAIL_CHANNEL TC,

       vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,

       vnadsprd.NAUT_TRAIL_ELEMENT TE,

       (SELECT VW.TRAIL_ID,

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

           AND TE.ELEMENT_TYPE = ''P''

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

                  ZSITE.CLLI) X

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

       ''NAUTILUS'' AS SOURCE,

       TC.CHANNEL_NAME,

       TC.SEQUENCE_NUMBER,

       TC.STATUS PORT_STATUS,

       TE.ELEMENT_TYPE,

       X.PROJECT_ID,

       X.ACLLI,

       X.ZCLLI,

       CURRENT_TIMESTAMP()

  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,

       vnadsprd.NAUT_TRAIL_CHANNEL TC,

        vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,

       vnadsprd.NAUT_TRAIL_ELEMENT TE,

       (SELECT VW.TRAIL_ID,

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

           AND TE.ELEMENT_TYPE = ''P''

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

                  ZSITE.CLLI) X

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

       ''NAUTILUS'' AS SOURCE,

       TC.CHANNEL_NAME,

       TC.SEQUENCE_NUMBER,

       TC.STATUS PORT_STATUS,

       TE.ELEMENT_TYPE,

       X.PROJECT_ID,

       X.ACLLI,

       X.ZCLLI,

       CURRENT_TIMESTAMP()

  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,

       vnadsprd.NAUT_TRAIL_CHANNEL TC,

        vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,

       vnadsprd.NAUT_TRAIL_ELEMENT TE,

       (SELECT VW.TRAIL_ID,

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

           AND TE.ELEMENT_TYPE = ''P''

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

                  ZSITE.CLLI) X

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

       ''NAUTILUS'' AS SOURCE,

       TC.CHANNEL_NAME,

       TC.SEQUENCE_NUMBER,

       TC.STATUS PORT_STATUS,

       TE.ELEMENT_TYPE,

       X.PROJECT_ID,

       X.ACLLI,

       X.ZCLLI,

       CURRENT_TIMESTAMP()

  FROM `JSON_TO_BQ.BQ_DATASET.vnadsprd`.NAUT_TRAIL T,

       vnadsprd.NAUT_TRAIL_CHANNEL TC,

        vnadsprd.NAUT_TRAIL_COMPONENT_ELEMENT TCE,

       vnadsprd.NAUT_TRAIL_ELEMENT TE,

       (SELECT VW.TRAIL_ID,

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

           AND TE.ELEMENT_TYPE = ''P''

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

                  ZSITE.CLLI) X

 WHERE T.TRAIL_ID = X.PARENT_TRAIL_ID

   AND TC.PARENT_TRAIL_ID = X.PARENT_TRAIL_ID

   AND TC.CHANNEL_NAME = to_char(X.CHANNEL_NAME)

    AND TE.CHANNEL_NAME= to_char(X.CHANNEL_NAME)

   AND TE.PARENT_TRAIL_ID= X.PARENT_TRAIL_ID

   AND TCE.ELEMENT_ID=TE.ELEMENT_ID'

-- BLOCK 3
-- complexity_score: 0
-- used_llm: False
COMMIT

-- BLOCK 4
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'TRUNCATE TABLE vnadsprd.IEN_MARKET_INPUT_LIST'

-- BLOCK 5
-- complexity_score: 4
-- used_llm: True
EXECUTE IMMEDIATE 'drop index vnadsprd.IEN_TRAIL_NAME_IDX'

-- BLOCK 6
-- complexity_score: 4
-- used_llm: True
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
-- note: Qualified simple table references.
INSERT INTO `JSON_TO_BQ.BQ_DATASET.vnadsprd`.IEN_MARKET_INPUT_LIST (
  TRAIL_ID,
  MARKET_NAME,
  MARKET_STATE,
  CLLI,
  TRAIL_NAME,
  TID_LOGICAL,
  ALTERNATE_NAME,
  NFID,
  CREATED_TS
)
SELECT DISTINCT
  G.TRAIL_ID,
  IFNULL(MARKET.VZB_MARKET_CITY_OR_VZT_REGION, 'Not Available') AS MARKET_NAME,
  CASE
    WHEN MARKET.STATE IS NOT NULL THEN MARKET.STATE
    ELSE IFNULL(SUBSTR(S.CLLI, 5, 2), 'N/A')
  END AS MARKET_STATE,
  IFNULL(S.CLLI, 'N/A') AS CLLI,
  G.TRAIL_NAME AS TRAIL_NAME,
  CASE
    WHEN E.SHELF_TYPE = 'MSERI' THEN MSE.TID_LOGICAL
    ELSE E.TID_LOGICAL
  END AS TID_LOGICAL,
  CASE
    WHEN E.SHELF_TYPE = 'MSERI' THEN MSE.ALTERNATE_NAME
    ELSE E.ALTERNATE_NAME
  END AS ALTERNATE_NAME,
  IFNULL(G.NF_ID, 'N/A') AS NFID,
  CURRENT_TIMESTAMP() AS CREATED_TS
FROM
  vnadsprd.IEN_MARKET_INPUT_LIST_STG G
  JOIN `JSON_TO_BQ.BQ_DATASET.vnadsprd`.icoe_pvnr_t_logical_port P
    ON P.PORT_REFERENCE_ID = G.PORT_REFERENCE_ID
  JOIN `JSON_TO_BQ.BQ_DATASET.vnadsprd`.ICOE_EQUIPMENT_TBL E
    ON P.EQP_REFERENCE_ID = E.EQP_REFERENCE_ID
  LEFT JOIN `JSON_TO_BQ

-- BLOCK 7
-- complexity_score: 0
-- used_llm: False
COMMIT

-- BLOCK 8
-- complexity_score: 4
-- used_llm: True
-- CREATE INDEX not supported in BigQuery; index creation omitted.

-- BLOCK 9
-- complexity_score: 0
-- used_llm: False
END
