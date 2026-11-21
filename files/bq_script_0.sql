-- BLOCK 1
-- complexity_score: 5
-- used_llm: True
CREATE OR REPLACE PROCEDURE DEMO_MARKET_DATA_LOAD()
AS
BEGIN
  TRUNCATE TABLE DEMO_MARKET_STG;
END;

-- BLOCK 2
-- complexity_score: 0
-- used_llm: False
-- note: Applied deterministic function mapping (NVL/SYSDATE/etc.).
-- note: Qualified simple table references.
INSERT INTO `JSON_TO_BQ.BQ_DATASET.DEMO_MARKET_STG` (
        PROJECT_ID,
        SITE_CODE,
        REGION,
        DEVICE_TYPE,
        PORT_NUMBER,
        CHANNEL_ID,
        CREATED_DATE
    )

    SELECT DISTINCT
        SRC.PROJECT_ID,
        SITE.SITE_CODE,
        SITE.REGION,
        DEV.DEVICE_TYPE,
        DEV.PORT_NUMBER,
        DEV.CHANNEL_ID,
        CURRENT_TIMESTAMP()

    FROM
        `JSON_TO_BQ.BQ_DATASET.DEMO_PROJECTS` SRC

        INNER JOIN `JSON_TO_BQ.BQ_DATASET.DEMO_SITES` SITE
            ON SRC.SITE_ID = SITE.SITE_ID

        LEFT JOIN `JSON_TO_BQ.BQ_DATASET.DEMO_DEVICES` DEV
            ON SITE.SITE_ID = DEV.SITE_ID

    WHERE
        SRC.STATUS = 'ACTIVE'

-- BLOCK 3
-- complexity_score: 0
-- used_llm: False
COMMIT

-- BLOCK 4
-- complexity_score: 0
-- used_llm: False
END DEMO_MARKET_DATA_LOAD

-- BLOCK 5
-- complexity_score: 0
-- used_llm: False
/
