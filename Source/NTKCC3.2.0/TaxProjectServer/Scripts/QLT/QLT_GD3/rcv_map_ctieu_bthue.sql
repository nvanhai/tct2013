CREATE TABLE rcv_map_ctieu_bthue
    (id                             NUMBER(10,0) NOT NULL,
    loai_tkhai                     VARCHAR2(2),
    ma_ctieu                       VARCHAR2(6),
    ma_muc                         VARCHAR2(3),
    ma_tmuc                        VARCHAR2(2))
/
-- Constraints for RCV_MAP_CTIEU_BTHUE
ALTER TABLE rcv_map_ctieu_bthue
ADD CONSTRAINT rcv_mciteu_bthue_pk PRIMARY KEY (id)
USING INDEX
/
