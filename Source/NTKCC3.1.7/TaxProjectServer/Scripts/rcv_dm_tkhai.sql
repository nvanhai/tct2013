-- Start of DDL Script for Table QLT_OWNER.RCV_DM_TKHAI
-- Generated 12-Dec-2005 11:26:41 from QLT_OWNER@QLT_93

CREATE TABLE rcv_dm_tkhai
    (ma                             VARCHAR2(2) NOT NULL,
    ten                            VARCHAR2(60) NOT NULL)
/

-- Create synonym RCV_DM_TKHAI
CREATE PUBLIC SYNONYM rcv_dm_tkhai
  FOR rcv_dm_tkhai
/

-- Grants for Table
GRANT DELETE ON rcv_dm_tkhai TO qlt
/
GRANT INSERT ON rcv_dm_tkhai TO qlt
/
GRANT SELECT ON rcv_dm_tkhai TO qlt
/
GRANT UPDATE ON rcv_dm_tkhai TO qlt
/
GRANT SELECT ON rcv_dm_tkhai TO qlt_read
/



-- Constraints for RCV_DM_TKHAI

ALTER TABLE rcv_dm_tkhai
ADD CONSTRAINT rcv_dtk_pk PRIMARY KEY (ma)
USING INDEX
/


-- End of DDL Script for Table QLT_OWNER.RCV_DM_TKHAI

