-- Start of DDL Script for View QLT_OWNER.RCV_V_DTNT
-- Generated 2/9/2006 4:11:19 PM from QLT_OWNER@QLT_91

CREATE OR REPLACE VIEW rcv_v_dtnt (
   tin,
   ten_dtnt,
   ma_cqt,
   ma_tinh,
   ma_huyen,
   ma_cap,
   ma_chuong,
   ma_loai,
   ma_khoan,
   dia_chi,
   trang_thai,
   ma_phong,
   ma_canbo,
   dien_thoai,
   fax,
   email,
   ngay_kdoanh,
   ngay_tchinh,
   loai )
AS
SELECT tp.tin tin,
       SUBSTR(tp.norm_name,1,60) ten_dtnt,
       tp.pay_taxo_id ma_cqt,
       tp.tran_prov ma_tinh,
       tp.tran_dist ma_huyen,
       tp.level_code ma_cap,
       tp.category ma_chuong,
       tp.group_code ma_loai,
       tp.chapter ma_khoan,
       tp.tran_addr dia_chi,
       tp.status trang_thai,
       tp.depa_id ma_phong,
       tp.staff_id ma_canbo,
       tp.tran_tel dien_thoai,
       tp.tran_fax fax,
       tp.tin_emal email,
       tp.start_date ngay_kdoanh,
       tp.fina_start_date ngay_tchinh,
       tp.payer_type loai
FROM tin_payer tp
WHERE (tp.update_no = 0)
  AND (tp.tin NOT IN (SELECT up.tin
                      FROM tin_unfrequent_payer up
					  WHERE (up.update_no = 0)
					    AND (up.status = '00')
						AND (tp.status <> '00'))
	  )
  AND (tp.tin NOT IN (SELECT pp.tin
                      FROM tin_personal_payer pp
					  WHERE (pp.update_no = 0)
					    AND (pp.status = '00')
						AND (tp.status <> '00'))
	  )
UNION ALL
SELECT pp.tin tin,
       SUBSTR(pp.norm_name,1,60) ten_dtnt,
       pp.pay_taxo_id ma_cqt,
       pp.tran_prov ma_tinh,
       pp.tran_dist ma_huyen,
       pp.level_code ma_cap,
       pp.category ma_chuong,
       pp.group_code ma_loai,
       pp.chapter ma_khoan,
       pp.tran_addr dia_chi,
       pp.status trang_thai,
       pp.depa_id ma_phong,
       pp.staff_id ma_canbo,
       pp.tran_tel dien_thoai,
       pp.tran_fax fax,
       pp.tin_emal email,
       TO_DATE(NULL) ngay_kdoanh,
       TO_DATE(NULL) ngay_tchinh,
       pp.payer_type loai
FROM tin_personal_payer pp
WHERE (pp.update_no = 0)
  AND (pp.tin NOT IN(
					SELECT tp.tin tin
					FROM tin_payer tp
					WHERE (tp.update_no = 0)
					  AND (tp.tin NOT IN (SELECT up.tin
					                      FROM tin_unfrequent_payer up
										  WHERE (up.update_no = 0)
										    AND (up.status = '00')
											AND (tp.status <> '00'))
						  )
					  AND (tp.tin NOT IN (SELECT pp.tin
					                      FROM tin_personal_payer pp
										  WHERE (pp.update_no = 0)
										    AND (pp.status = '00')
											AND (tp.status <> '00'))
						  )
					)
      )
UNION ALL
SELECT up.tin tin,
       SUBSTR(up.norm_name,1,60) ten_dtnt,
       up.pay_taxo_id ma_cqt,
       up.tran_prov ma_tinh,
       up.tran_dist ma_huyen,
       up.level_code ma_cap,
       up.category ma_chuong,
       up.group_code ma_loai,
       up.chapter ma_khoan,
       up.tran_addr dia_chi,
       up.status trang_thai,
       up.depa_id ma_phong,
       up.staff_id ma_canbo,
       up.tran_tel dien_thoai,
       up.tran_fax fax,
       up.tin_emal email,
       up.start_date ngay_kdoanh,
       up.fina_start_date ngay_tchinh,
       up.payer_type loai
FROM tin_unfrequent_payer up
WHERE (up.update_no = 0)
  AND (up.tin NOT IN(
					SELECT tp.tin tin
					FROM tin_payer tp
					WHERE (tp.update_no = 0)
					  AND (tp.tin NOT IN (SELECT up.tin
					                      FROM tin_unfrequent_payer up
										  WHERE (up.update_no = 0)
										    AND (up.status = '00')
											AND (tp.status <> '00'))
						  )
					  AND (tp.tin NOT IN (SELECT pp.tin
					                      FROM tin_personal_payer pp
										  WHERE (pp.update_no = 0)
										    AND (pp.status = '00')
											AND (tp.status <> '00'))
						  )
					)
      )
  AND (up.tin NOT IN(
                    SELECT pp.tin tin
                    FROM tin_personal_payer pp
                    WHERE (pp.update_no = 0)
                      AND (pp.tin NOT IN(
                    					SELECT tp.tin tin
                    					FROM tin_payer tp
                    					WHERE (tp.update_no = 0)
                    					  AND (tp.tin NOT IN (SELECT up.tin
                    					                      FROM tin_unfrequent_payer up
                    										  WHERE (up.update_no = 0)
                    										    AND (up.status = '00')
                    											AND (tp.status <> '00'))
                    						  )
                    					  AND (tp.tin NOT IN (SELECT pp.tin
                    					                      FROM tin_personal_payer pp
                    										  WHERE (pp.update_no = 0)
                    										    AND (pp.status = '00')
                    											AND (tp.status <> '00'))
                    						  )
                    					)
                          )
                    )
      )
/

-- End of DDL Script for View QLT_OWNER.RCV_V_DTNT

