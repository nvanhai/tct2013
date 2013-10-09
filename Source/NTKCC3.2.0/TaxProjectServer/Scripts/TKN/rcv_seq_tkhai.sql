-- Start of DDL Script for Sequence TKN_OWNER.RCV_SEQ_TKHAI
-- Generated 8-Dec-2005 18:54:22 from TKN_OWNER@TKN_93

CREATE SEQUENCE rcv_seq_tkhai
  INCREMENT BY 1
  START WITH 1
  MINVALUE 1
  MAXVALUE 999999999999999999999999999
  NOCYCLE
  NOORDER
  NOCACHE
/

CREATE PUBLIC SYNONYM rcv_seq_tkhai
  FOR rcv_seq_tkhai
/
  
-- Grants for Sequence
GRANT SELECT ON rcv_seq_tkhai TO tkn
/
GRANT SELECT ON rcv_seq_tkhai TO tkn_read
/
-- End of DDL Script for Sequence TKN_OWNER.RCV_SEQ_TKHAI

