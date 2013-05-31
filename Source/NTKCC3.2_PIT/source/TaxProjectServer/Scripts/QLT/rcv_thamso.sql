CREATE TABLE rcv_thamso
    (ten                            VARCHAR2(50) NOT NULL,
    gia_tri                        VARCHAR2(50),
    ghi_chu                        VARCHAR2(100))
/
-- Comments for RCV_THAMSO
COMMENT ON COLUMN rcv_thamso.ghi_chu IS 'Giai thich cho tham so'
/
COMMENT ON COLUMN rcv_thamso.gia_tri IS 'Gia tri cua tham so'
/
COMMENT ON COLUMN rcv_thamso.ten IS 'Ten tham so'
/

