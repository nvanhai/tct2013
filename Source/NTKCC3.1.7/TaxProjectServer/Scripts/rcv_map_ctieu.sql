-- Start of DDL Script for Table QLT_OWNER.RCV_MAP_CTIEU
-- Generated 12-Dec-2005 11:29:04 from QLT_OWNER@QLT_93

CREATE TABLE rcv_map_ctieu
    (loai_dlieu                     VARCHAR2(4) NOT NULL,
    ky_hieu                        VARCHAR2(10) NOT NULL,
    kieu_dlieu                     VARCHAR2(1) NOT NULL,
    gdn_id                         NUMBER(10,0) NOT NULL,
    ky_hieu_ctieu                  VARCHAR2(4))
/

-- Create synonym RCV_MAP_CTIEU
CREATE PUBLIC SYNONYM rcv_map_ctieu
  FOR rcv_map_ctieu
/

-- Grants for Table
GRANT DELETE ON rcv_map_ctieu TO qlt
/
GRANT INSERT ON rcv_map_ctieu TO qlt
/
GRANT SELECT ON rcv_map_ctieu TO qlt
/
GRANT UPDATE ON rcv_map_ctieu TO qlt
/
GRANT SELECT ON rcv_map_ctieu TO qlt_read
/



-- Constraints for RCV_MAP_CTIEU


ALTER TABLE rcv_map_ctieu
ADD CONSTRAINT rcv_mctieu_uk UNIQUE (loai_dlieu, ky_hieu)
USING INDEX
/


-- End of DDL Script for Table QLT_OWNER.RCV_MAP_CTIEU

-- Foreign Key
ALTER TABLE rcv_map_ctieu
ADD CONSTRAINT rcv_mctieu_fk FOREIGN KEY (gdn_id)
REFERENCES rcv_gdien_tkhai (id)
/
-- End of DDL script for Foreign Key(s)
