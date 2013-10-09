CREATE TABLE rcv_tkhai_hdr
    (id                             NUMBER(10,0) NOT NULL,
    tin                            VARCHAR2(14) NOT NULL,
    ten_dtnt                       VARCHAR2(100) NOT NULL,
    dia_chi                        VARCHAR2(60),
    loai_tkhai                     VARCHAR2(2),
    ngay_nop                       DATE,
    kylb_tu_ngay                   DATE,
    kylb_den_ngay                  DATE,
    kykk_tu_ngay                   DATE,
    kykk_den_ngay                  DATE,
    ngay_cap_nhat                  DATE,
    nguoi_cap_nhat                 VARCHAR2(60),
    co_loi_ddanh                   CHAR(1),
    so_hieu_tep                    VARCHAR2(20),
    so_tt_tk                       NUMBER(10,0),
    da_nhan                        CHAR(1),
    ghi_chu_loi                    VARCHAR2(100),
    co_gtrinh_02a                  CHAR(1),
    co_gtrinh_02b                  CHAR(1),
    co_gtrinh_02c                  CHAR(1),
    khoa_so                        VARCHAR2(1),
    tu_ngay                        DATE,
    den_ngay                       DATE,
    phong_xly                      VARCHAR2(7))
/
-- Constraints for RCV_TKHAI_HDR
ALTER TABLE rcv_tkhai_hdr
ADD CONSTRAINT rcv_tkh_pk PRIMARY KEY (id)
USING INDEX
/
-- Triggers for RCV_TKHAI_HDR
CREATE OR REPLACE TRIGGER rcv_trg_auid_tkhai_hdr
 BEFORE
  INSERT OR DELETE OR UPDATE
 ON rcv_tkhai_hdr
REFERENCING NEW AS NEW OLD AS OLD
 FOR EACH ROW
DECLARE
    CURSOR c_Exist_Tep_Hoso(p_So_Hieu VARCHAR2, 
                            p_Loai_Tkhai VARCHAR2) IS SELECT COUNT(so_hieu_tep) so_tkhai FROM rcv_tkhai_hdr
                                                      WHERE so_hieu_tep = p_So_Hieu
                                                      AND Loai_tkhai = p_loai_tkhai;
                               
    vc_Exist_Tep_Hoso c_Exist_Tep_Hoso%ROWTYPE;
BEGIN
    IF INSERTING THEN
        OPEN c_Exist_Tep_Hoso(:NEW.so_hieu_tep, :NEW.loai_tkhai);
        FETCH c_Exist_Tep_Hoso INTO vc_Exist_Tep_Hoso;
        IF c_Exist_Tep_Hoso%FOUND THEN
            :NEW.so_tt_tk := vc_Exist_Tep_Hoso.so_tkhai + 1;
        ELSE
            :NEW.so_tt_tk := NVL(:OLD.so_tt_tk,1);
        END IF;
        CLOSE c_Exist_Tep_Hoso;    
    END IF;    
END;
/
-- Foreign Key
ALTER TABLE rcv_tkhai_hdr
ADD CONSTRAINT rcv_tkh_fk FOREIGN KEY (loai_tkhai)
REFERENCES rcv_dm_tkhai (ma)
/
