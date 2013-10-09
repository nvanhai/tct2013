CREATE TABLE rcv_tkhai_dtl
    (id                             NUMBER(10,0) NOT NULL,
    hdr_id                         NUMBER(10,0),
    loai_dlieu                     VARCHAR2(4) NOT NULL,
    ky_hieu                        VARCHAR2(10) NOT NULL,
    gia_tri                        VARCHAR2(1000),
    row_id                         NUMBER(10,0))
/
-- Indexes for RCV_TKHAI_DTL
CREATE INDEX rcv_tkhai_dtl_ind ON rcv_tkhai_dtl
  (
    hdr_id                          ASC
  )
/
-- Constraints for RCV_TKHAI_DTL
ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_pk PRIMARY KEY (id)
USING INDEX
/
-- Foreign Key
ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_hdr_fk FOREIGN KEY (hdr_id)
REFERENCES rcv_tkhai_hdr (id) ON DELETE CASCADE
/
ALTER TABLE rcv_tkhai_dtl
ADD CONSTRAINT rcv_tkd_mctieu_fk FOREIGN KEY (loai_dlieu, ky_hieu)
REFERENCES rcv_map_ctieu (loai_dlieu,ky_hieu)
/
