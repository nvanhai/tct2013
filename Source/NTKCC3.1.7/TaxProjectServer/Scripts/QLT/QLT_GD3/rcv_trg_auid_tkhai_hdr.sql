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

