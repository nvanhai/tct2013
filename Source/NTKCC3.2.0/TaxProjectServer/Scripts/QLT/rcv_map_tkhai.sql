CREATE TABLE rcv_map_tkhai
    (ma_tkhai                       VARCHAR2(2) NOT NULL,
    ma_tkhai_qlt                   VARCHAR2(2) NOT NULL,
    ghi_chu                        VARCHAR2(100),
    loai                           VARCHAR2(2) NOT NULL,
    nhom_hso                       VARCHAR2(2))
/
-- Constraints for RCV_MAP_TKHAI
ALTER TABLE rcv_map_tkhai
ADD CONSTRAINT rcv_map_tkhai_pk PRIMARY KEY (ma_tkhai)
USING INDEX
/

