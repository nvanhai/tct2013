-- Start of DDL Script for Table QLT_OWNER.RCV_TKHAI_HDR
-- Generated 12-Dec-2005 11:32:07 from QLT_OWNER@QLT_93

CREATE TABLE rcv_tkhai_hdr
    (id                             NUMBER(10,0) NOT NULL,
    tin                            VARCHAR2(14) NOT NULL,
    ten_dtnt                       VARCHAR2(100) NOT NULL,
    dia_chi                        VARCHAR2(60),
    loai_tkhai                     VARCHAR2(2) NOT NULL,
    ngay_nop                       DATE NOT NULL,
    kylb_tu_ngay                   DATE NOT NULL,
    kylb_den_ngay                  DATE NOT NULL,
    kykk_tu_ngay                   DATE NOT NULL,
    kykk_den_ngay                  DATE NOT NULL,
    ngay_cap_nhat                  DATE NOT NULL,
    nguoi_cap_nhat                 VARCHAR2(60) NOT NULL,
    co_loi_ddanh                   CHAR(1),
    so_hieu_tep                    VARCHAR2(20),
    so_tt_tk                       NUMBER(10,0),
    da_nhan                        CHAR(1),
    ghi_chu_loi                    VARCHAR2(100),
    co_gtrinh_02a                  CHAR(1),
    co_gtrinh_02b                  CHAR(1),
    co_gtrinh_02c                  CHAR(1))
/

-- Create synonym RCV_TKHAI_HDR
CREATE PUBLIC SYNONYM rcv_tkhai_hdr
  FOR rcv_tkhai_hdr
/

-- Grants for Table
GRANT DELETE ON rcv_tkhai_hdr TO qlt
/
GRANT INSERT ON rcv_tkhai_hdr TO qlt
/
GRANT SELECT ON rcv_tkhai_hdr TO qlt
/
GRANT UPDATE ON rcv_tkhai_hdr TO qlt
/
GRANT SELECT ON rcv_tkhai_hdr TO qlt_read
/



-- Constraints for RCV_TKHAI_HDR


ALTER TABLE rcv_tkhai_hdr
ADD CONSTRAINT rcv_tkh_pk PRIMARY KEY (id)
USING INDEX
/


-- End of DDL Script for Table QLT_OWNER.RCV_TKHAI_HDR

-- Foreign Key
ALTER TABLE rcv_tkhai_hdr
ADD CONSTRAINT rcv_tkh_fk FOREIGN KEY (loai_tkhai)
REFERENCES rcv_dm_tkhai (ma)
/
-- End of DDL script for Foreign Key(s)
