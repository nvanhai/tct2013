CREATE TABLE rcv_map_ctieu
    (loai_dlieu                     VARCHAR2(4) NOT NULL,
    ky_hieu                        VARCHAR2(10) NOT NULL,
    kieu_dlieu                     VARCHAR2(1),
    gdn_id                         NUMBER(10,0) NOT NULL,
    ky_hieu_ctieu                  VARCHAR2(4))
/
-- Constraints for RCV_MAP_CTIEU
ALTER TABLE rcv_map_ctieu
ADD CONSTRAINT rcv_mctieu_uk UNIQUE (loai_dlieu, ky_hieu)
USING INDEX
/
-- Foreign Key
ALTER TABLE rcv_map_ctieu
ADD CONSTRAINT rcv_mctieu_fk FOREIGN KEY (gdn_id)
REFERENCES rcv_gdien_tkhai (id)
/
