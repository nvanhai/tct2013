-- Start of DDL Script for Sequence QLT_OWNER.RCV_XLTK_DTL_SEQ
-- Generated 8-Dec-2005 18:52:15 from QLT_OWNER@QLT_93

CREATE SEQUENCE rcv_xltk_dtl_seq
  INCREMENT BY 1
  START WITH 1
  MINVALUE 1
  MAXVALUE 999999999999999999999999999
  NOCYCLE
  NOORDER
  CACHE 20
/

-- Grants for Sequence
GRANT SELECT ON rcv_xltk_dtl_seq TO qlt
/
GRANT SELECT ON rcv_xltk_dtl_seq TO qlt_read
/

-- End of DDL Script for Sequence QLT_OWNER.RCV_XLTK_DTL_SEQ

