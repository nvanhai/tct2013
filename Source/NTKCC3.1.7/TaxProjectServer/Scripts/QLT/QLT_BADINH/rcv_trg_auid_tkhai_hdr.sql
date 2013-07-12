CREATE OR REPLACE TRIGGER rcv_trg_auid_tkhai_hdr
 BEFORE
  INSERT OR DELETE OR UPDATE
 ON rcv_tkhai_hdr
REFERENCING NEW AS NEW OLD AS OLD
 FOR EACH ROW
BEGIN
    IF INSERTING THEN
        :NEW.so_tt_tk := NVL(:OLD.so_tt_tk,0);
    END IF;
END;
/

