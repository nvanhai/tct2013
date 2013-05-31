-- Start of DDL Script for Table QLT_OWNER.RCV_TKHAI_DTL
-- Generated 12-Dec-2005 11:32:55 from QLT_OWNER@QLT_93

CREATE TABLE rcv_tkhai_dtl
    (id                             NUMBER(10,0) NOT NULL,
    hdr_id                         NUMBER(10,0) NOT NULL,
    loai_dlieu                     VARCHAR2(4) NOT NULL,
    ky_hieu                        VARCHAR2(10) NOT NULL,
    gia_tri                        VARCHAR2(1000),
    row_id                         NUMBER(10,0))
/

-- Create synonym RCV_TKHAI_DTL
CREATE PUBLIC SYNONYM rcv_tkhai_dtl
  FOR rcv_tkhai_dtl
/

-- Grants for Table
GRANT DELETE ON rcv_tkhai_dtl TO qlt
/
GRANT INSERT ON rcv_tkhai_dtl TO qlt
/
GRANT SELECT ON rcv_tkhai_dtl TO qlt
/
GRANT UPDATE ON rcv_tkhai_dtl TO qlt
/
GRANT SELECT ON rcv_tkhai_dtl TO qlt_read
/



-- Constraints for RCV_TKHAI_DTL



ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_pk PRIMARY KEY (id)
USING INDEX
/


-- End of DDL Script for Table QLT_OWNER.RCV_TKHAI_DTL

-- Foreign Key
ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_hdr_fk FOREIGN KEY (hdr_id)
REFERENCES rcv_tkhai_hdr (id) ON DELETE CASCADE
/
ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_mctieu_fk FOREIGN KEY (loai_dlieu, ky_hieu)
REFERENCES rcv_map_ctieu (loai_dlieu,ky_hieu)
/
-- End of DDL script for Foreign Key(s)
