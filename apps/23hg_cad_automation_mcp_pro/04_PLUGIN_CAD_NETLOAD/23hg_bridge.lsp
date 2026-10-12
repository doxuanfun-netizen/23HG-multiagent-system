;;; 23HG CAD AUTOMATION BRIDGE (AutoLISP Helper)
;;; Tác giả: Nguyễn Bảo Tú (@baotuhg) - 23HG SYSTEM
;;; Hỗ trợ các phím tắt và thao tác tương tác nhanh trong AutoCAD

(defun c:AI_DRAW ( / f line)
  (setq f (open (strcat (getenv "USERPROFILE") "\\cad_cmd.scr") "r"))
  (if f
    (progn
      (close f)
      (command "_SCRIPT" (strcat (getenv "USERPROFILE") "\\cad_cmd.scr"))
    )
    (princ "\n[23HG] Chua co file kich ban lenh tu AI.")
  )
  (princ)
)

(defun c:23HG_STATUS ()
  (princ "\n[23HG] AutoCAD da san sang ket noi truc tiep qua Native COM Interop!")
  (princ)
)

(princ "\n[23HG CAD AUTOMATION] Da nap thanh cong Bridge LISP. Go 23HG_STATUS de kiem tra.")
(princ)
