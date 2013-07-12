-- Start of DDL Script for Table QLT_OWNER.RCV_MAP_TKHAI
-- Generated 12-Dec-2005 11:30:10 from QLT_OWNER@QLT_93

CREATE TABLE rcv_map_tkhai
    (ma_tkhai                       VARCHAR2(2) NOT NULL,
    ma_tkhai_qlt                   VARCHAR2(2) NOT NULL,
    ghi_chu                        VARCHAR2(100))
/

-- Create synonym RCV_MAP_TKHAI
CREATE PUBLIC SYNONYM rcv_map_tkhai
  FOR rcv_map_tkhai
/

-- Grants for Table
GRANT DELETE ON rcv_map_tkhai TO qlt
/
GRANT INSERT ON rcv_map_tkhai TO qlt
/
GRANT SELECT ON rcv_map_tkhai TO qlt
/
GRANT UPDATE ON rcv_map_tkhai TO qlt
/
GRANT SELECT ON rcv_map_tkhai TO qlt_read
/



-- Constraints for RCV_MAP_TKHAI

ALTER TABLE rcv_map_tkhai
ADD CONSTRAINT rcv_map_tkhai_pk PRIMARY KEY (ma_tkhai)
USING INDEX
/


-- End of DDL Script for Table QLT_OWNER.RCV_MAP_TKHAI

