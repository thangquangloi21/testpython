import paramiko

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect('tvc-sv15', username='mfg', password='tvcadmin')

cmd = '''
export DLC=/qad/oe117
cd /qad/test/databases
$DLC/bin/_progres -db tvctest << 'EOF'
OUTPUT TO "/home/mfg/lad_det_export.csv".
PUT UNFORMATTED "lad_nbr,lad_part,lad_lot,lad_loc,lad_qty_all" SKIP.
FOR EACH lad_det NO-LOCK:
PUT UNFORMATTED 
STRING(lad_det.lad_nbr) + "," + 
STRING(lad_det.lad_part) + "," + 
STRING(lad_det.lad_lot) + "," + 
STRING(lad_det.lad_loc) + "," + 
STRING(lad_det.lad_qty_all)SKIP.
END.
OUTPUT CLOSE.


OUTPUT TO "/home/mfg/wod_det_export.csv".
PUT UNFORMATTED "wod_lot,wod_part,wod_qty_req,wod_qty_iss" SKIP.
FOR EACH wod_det NO-LOCK:
PUT UNFORMATTED 
STRING(wod_det.wod_lot) + "," + 
STRING(wod_det.wod_part) + "," + 
STRING(wod_det.wod_qty_req) + "," + 
STRING(wod_det.wod_qty_iss) SKIP.
END.
OUTPUT CLOSE.


OUTPUT TO "/home/mfg/wo_mstr_export.csv".
PUT UNFORMATTED "wo_lot,wo_nbr,wo__chr02,wo_part,wo_lot_next,wo_due_date" SKIP.
FOR EACH wo_mstr NO-LOCK where wo_mstr.wo_due_date >= DATE("01/01/25"):
PUT UNFORMATTED 
STRING(wo_mstr.wo_lot) + "," + 
STRING(wo_mstr.wo_nbr) + "," + 
STRING(wo_mstr.wo__chr02) + "," + 
STRING(wo_mstr.wo_part) + "," + 
STRING(wo_mstr.wo_lot_next) + "," + 
STRING(wo_mstr.wo_due_date, "99/99/99")SKIP.
END.
OUTPUT CLOSE.


OUTPUT TO "/home/mfg/pt_mstr_export.csv".
PUT UNFORMATTED "pt_part,pt_um,pt_prod_line,pt_part_type" SKIP.
FOR EACH pt_mstr NO-LOCK:
PUT UNFORMATTED 
STRING(pt_mstr.pt_part) + "," + 
STRING(pt_mstr.pt_um) + "," + 
STRING(pt_mstr.pt_prod_line) + "," + 
STRING(pt_mstr.pt_part_type)SKIP.
END.
OUTPUT CLOSE.

EOF
'''

stdin, stdout, stderr = ssh.exec_command(cmd)
print(stderr.read().decode())
sftp = ssh.open_sftp()
sftp.get('/home/mfg/lad_det_export.csv', 'D:/4.DEV/TEST_python/lad_det_export.csv')
sftp.get('/home/mfg/wod_det_export.csv', 'D:/4.DEV/TEST_python/wod_det_export.csv')
sftp.get('/home/mfg/wo_mstr_export.csv', 'D:/4.DEV/TEST_python/wo_mstr_export.csv')
sftp.get('/home/mfg/pt_mstr_export.csv', 'D:/4.DEV/TEST_python/pt_mstr_export.csv')
sftp.close()
ssh.close()

print("✓ File downloaded!")