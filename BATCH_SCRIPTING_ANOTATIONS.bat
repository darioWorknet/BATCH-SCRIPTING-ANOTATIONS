@echo off

:: OPENING FOLDER
	%SystemRoot%\explorer.exe "C:\Users\%USERNAME%\Desktop\M_y__F_o_l_d_e_r"


:: OPENING SPECIFIED EXCEL FILE
	SET params=%*
	START excel "C:\Users\%USERNAME%\Desktop\f_i_l_e_n_a_m_e.xlsm" /e/%params% 

        :: a_n_o_t_h_e_r_____o_p_t_i_o_n_____t_o_____o_p_e_n_____f_i_l_e_____w_i_t_h_____E_x_c_e_l
		::SET myProgram="C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"
		::SET myFile="C:\Users\%USERNAME%\Desktop\f_i_l_e_n_a_m_e.xlsm"
		::START   ""   %myProgram%    %myFile%
       


:: 1 SECOND DELAY (actually not needed)
	:: ping is a system utility that sends ping requests. ping is available on all versions of Windows.
	:: 127.0.0.1 is the IP address of localhost. This IP address is guaranteed to always resolve, be reachable, and immediately respond to pings
	:: -n 6 specifies that there are to be 6 pings. There is a 1s delay between each ping, so for a 5s delay you need to send 6 pings
	:: > nul suppress the output of ping, by redirecting it to nul
		ping 127.0.0.1 -n 2 > nul


:: OPENING EM_CLIENT
	START   "EmClient"     /D "C:\Program Files (x86)\eM Client"     MailClient.exe /WAIT /MIN


:: OPENING WHATSAPP
	START   "WhatsApp"     /D C:\Users\%USERNAME%\AppData\Local\WhatsApp     WhatsApp.exe /WAIT /MIN


::Move files from folder_A to folder_B
	SET origin=C:\Users\%USERNAME%\Desktop\F_o_l_d_e_r__A
	SET destination=C:\Users\%USERNAME%\Desktop\F_o_l_d_e_r__B
	MOVE %origin%\*.* %destination%

