FasdUAS 1.101.10   ��   ��    k             l     ����  r       	  m      
 
 �   2 A r c h i v e   s e l e c t e d   m e s s a g e s 	 o      ���� 0 myname myName��  ��        l    ����  r        m       �   " M i c r o s o f t   O u t l o o k  o      ���� 0 mailname mailName��  ��        l    ����  r        m    	   �   
 I n b o x  o      ���� 0 	inboxname 	inboxName��  ��        l    ����  r        m       �      A r c h i v e  o      ���� 0 archivename archiveName��  ��     ! " ! l     ��������  ��  ��   "  # $ # l  z %���� % O   z & ' & k   y ( (  ) * ) r     + , + 4   �� -
�� 
cwin - m    ����  , o      ���� 0 frontwin frontWin *  . / . r      0 1 0 n     2 3 2 1    ��
�� 
pnam 3 o    ���� 0 frontwin frontWin 1 o      ���� 0 winname winName /  4 5 4 r   ! & 6 7 6 1   ! $��
�� 
CMgs 7 o      ���� 0 currmsgs currMsgs 5  8 9 8 Z   ' B : ;���� : =   ' + < = < o   ' (���� 0 currmsgs currMsgs = J   ( *����   ; k   . > > >  ? @ ? I  . ;�� A B
�� .sysodisAaleR        TEXT A l  . 1 C���� C b   . 1 D E D m   . / F F � G G @ N o   s e l e c t e d   m e s s a g e s   i n   w i n d o w :   E o   / 0���� 0 winname winName��  ��   B �� H��
�� 
mesS H m   4 7 I I � J J ( N o   m e s s a g e s   s e l e c t e d��   @  K�� K L   < > L L m   < =����  ��  ��  ��   9  M N M r   C M O P O n   C I Q R Q 4   D I�� S
�� 
cobj S m   G H����  R o   C D���� 0 currmsgs currMsgs P o      ���� 0 firstmsg firstMsg N  T U T r   N W V W V 1   N S��
�� 
pOMC W o      ���� 0 onmycomputer onMyComputer U  X Y X l  X X��������  ��  ��   Y  Z [ Z l  X X�� \ ]��   \ C = Point to archive folders (Archive/Received and Archive/Sent)    ] � ^ ^ z   P o i n t   t o   a r c h i v e   f o l d e r s   ( A r c h i v e / R e c e i v e d   a n d   A r c h i v e / S e n t ) [  _ ` _ Q   X � a b c a k   [ � d d  e f e r   [ g g h g n   [ c i j i 4   ^ c�� k
�� 
cFld k o   a b���� 0 archivename archiveName j o   [ ^���� 0 onmycomputer onMyComputer h o      ���� 0 archivefolder archiveFolder f  l m l r   h v n o n n   h r p q p 4   k r�� r
�� 
cFld r m   n q s s � t t  S e n t q o   h k���� 0 archivefolder archiveFolder o o      ����  0 destsentfolder destSentFolder m  u�� u r   w � v w v n   w � x y x 4   z ��� z
�� 
cFld z m   } � { { � | |  R e c e i v e d y o   w z���� 0 archivefolder archiveFolder w o      ����  0 destrecvfolder destRecvFolder��   b R      �� } ~
�� .ascrerr ****      � **** } o      ���� 0 errormessage errorMessage ~ �� ��
�� 
errn  o      ���� 0 errornumber errorNumber��   c k   � � � �  � � � I  � ��� � �
�� .sysodisAaleR        TEXT � l  � � ����� � m   � � � � � � � 0 A r c h i v e   f o l d e r   n o t   f o u n d��  ��   � �� � �
�� 
mesS � l  � � ����� � b   � � � � � b   � � � � � o   � ����� 0 errormessage errorMessage � m   � � � � � � �  E r r o r   n u m b e r :   � o   � ����� 0 errornumber errorNumber��  ��   � �� ���
�� 
as A � m   � ���
�� EAlTcriT��   �  ��� � L   � � � � m   � �����  ��   `  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � %  Count messages and notify user    � � � � >   C o u n t   m e s s a g e s   a n d   n o t i f y   u s e r �  � � � r   � � � � � l  � � ����� � I  � ��� ���
�� .corecnte****       **** � n  � � � � � 2  � ���
�� 
cobj � o   � ����� 0 currmsgs currMsgs��  ��  ��   � o      ���� 0 msgcount msgCount �  � � � I  � ����� �
�� .sysonotfnull��� ��� TEXT��   � �� � �
�� 
appr � o   � ����� 0 myname myName � �� ���
�� 
subt � l  � � ����� � b   � � � � � b   � � � � � b   � � � � � m   � � � � � � � & A t t e m p t i n g   t o   m o v e   � o   � ����� 0 msgcount msgCount � m   � � � � � � �    m e s s a g e s   t o   � o   � ����� 0 archivename archiveName��  ��  ��   �  � � � l  � ���������  ��  ��   �  � � � l  � ��� � ���   � X R Iterate over selected messages and archive based on whether they're sent/received    � � � � �   I t e r a t e   o v e r   s e l e c t e d   m e s s a g e s   a n d   a r c h i v e   b a s e d   o n   w h e t h e r   t h e y ' r e   s e n t / r e c e i v e d �  � � � l  � ��� � ���   � - ' and use the date to file messages away    � � � � N   a n d   u s e   t h e   d a t e   t o   f i l e   m e s s a g e s   a w a y �  � � � r   � � � � � 1   � ���
�� 
dfAc � o      ���� 0 
defaccount 
defAccount �  � � � r   � � � � � n   � � � � � m   � ���
�� 
emad � o   � ����� 0 
defaccount 
defAccount � o      ����  0 defsenderemail defSenderEmail �  � � � Q   �_ � � � � X   �; ��� � � k  6 � �  � � � r  
 � � � n   � � � 1  ��
�� 
sndr � o  ���� 0 
themessage 
theMessage � o      ���� 0 	senderobj 	senderObj �  � � � r   � � � n   � � � 1  ��
�� 
radd � o  ���� 0 	senderobj 	senderObj � o      ���� 0 senderemail senderEmail �  ��� � Z  6 � ��� � � l  ����� � =   � � � o  ���� 0 senderemail senderEmail � o  ����  0 defsenderemail defSenderEmail��  ��   � n !* � � � I  "*�� �����  0 archivemessage archiveMessage �  � � � o  "#���� 0 
themessage 
theMessage �  ��� � o  #&����  0 destsentfolder destSentFolder��  ��   �  f  !"��   � n -6 � � � I  .6�� �����  0 archivemessage archiveMessage �  � � � o  ./�� 0 
themessage 
theMessage �  ��~ � o  /2�}�}  0 destrecvfolder destRecvFolder�~  ��   �  f  -.��  �� 0 
themessage 
theMessage � o   � ��|�| 0 currmsgs currMsgs � R      �{ � �
�{ .ascrerr ****      � **** � o      �z�z 0 errormessage errorMessage � �y ��x
�y 
errn � o      �w�w 0 errornumber errorNumber�x   � k  C_ � �  � � � I C\�v � �
�v .sysodisAaleR        TEXT � m  CF � � � � �   A r c h i v i n g   f a i l e d � �u � �
�u 
mesS � l IP ��t�s � b  IP � � � b  IN   o  IJ�r�r 0 errormessage errorMessage m  JM �  
 E r r o r   n u m b e r :   � o  NO�q�q 0 errornumber errorNumber�t  �s   � �p�o
�p 
as A m  SV�n
�n EAlTcriT�o   � �m L  ]_ m  ]^�l�l  �m   � �k I `y�j�i
�j .sysonotfnull��� ��� TEXT�i   �h	

�h 
appr	 o  de�g�g 0 myname myName
 �f�e
�f 
subt l hs�d�c b  hs b  ho m  hk � , S u c c e s s f u l l y   a r c h i v e d   o  kn�b�b 0 msgcount msgCount m  or �    m e s s a g e s�d  �c  �e  �k   ' m    �                                                                                  OPIM  alis    D  gala-gs                        BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    g a l a - g s  "Applications/Microsoft Outlook.app  / ��  ��  ��   $  l     �a�`�_�a  �`  �_    l     �^�^   ^ X Subroutine to archive a message to a specific destination folder, creating a sub-folder    � �   S u b r o u t i n e   t o   a r c h i v e   a   m e s s a g e   t o   a   s p e c i f i c   d e s t i n a t i o n   f o l d e r ,   c r e a t i n g   a   s u b - f o l d e r  l     �] �]   _ Y based on the year of the message. Checks to see if the folder exists prior to moving it.     �!! �   b a s e d   o n   t h e   y e a r   o f   t h e   m e s s a g e .   C h e c k s   t o   s e e   i f   t h e   f o l d e r   e x i s t s   p r i o r   t o   m o v i n g   i t . "#" i     $%$ I      �\&�[�\  0 archivemessage archiveMessage& '(' o      �Z�Z 0 
themessage 
theMessage( )�Y) o      �X�X 0 
destfolder 
destFolder�Y  �[  % O     {*+* k    z,, -.- r    	/0/ m    �W
�W boovtrue0 n      121 l   3�V�U3 1    �T
�T 
pRed�V  �U  2 o    �S�S 0 
themessage 
theMessage. 454 r   
 676 n   
 898 1    �R
�R 
rTim9 o   
 �Q�Q 0 
themessage 
theMessage7 o      �P�P 0 daterecv dateRecv5 :;: r    <=< n    >?> 1    �O
�O 
year? o    �N�N 0 daterecv dateRecv= o      �M�M 0 yearrecv yearRecv; @A@ r    BCB l   D�L�KD b    EFE m    GG �HH  F o    �J�J 0 yearrecv yearRecv�L  �K  C o      �I�I 0 yearrecvstr yearRecvStrA IJI l   �H�G�F�H  �G  �F  J K�EK Z    zLM�DNL l   $O�C�BO I   $�AP�@
�A .coredoexnull���     obj P n     QRQ 4     �?S
�? 
cFldS o    �>�> 0 yearrecvstr yearRecvStrR o    �=�= 0 
destfolder 
destFolder�@  �C  �B  M k   ' 5TT UVU r   ' -WXW n   ' +YZY 4   ( +�<[
�< 
cFld[ o   ) *�;�; 0 yearrecvstr yearRecvStrZ o   ' (�:�: 0 
destfolder 
destFolderX o      �9�9 0 	thefolder 	theFolderV \�8\ I  . 5�7]^
�7 .coremovenull���     obj ] o   . /�6�6 0 
themessage 
theMessage^ �5_�4
�5 
insh_ o   0 1�3�3 0 	thefolder 	theFolder�4  �8  �D  N k   8 z`` aba Q   8 kcdec I  ; I�2�1f
�2 .corecrel****      � null�1  f �0gh
�0 
koclg m   = >�/
�/ 
cMFoh �.ij
�. 
inshi o   ? @�-�- 0 
destfolder 
destFolderj �,k�+
�, 
prdtk K   A Ell �*m�)
�* 
pnamm o   B C�(�( 0 yearrecvstr yearRecvStr�)  �+  d R      �'no
�' .ascrerr ****      � ****n o      �&�& 0 errormessage errorMessageo �%p�$
�% 
errnp o      �#�# 0 errornumber errorNumber�$  e k   Q kqq rsr I  Q h�"tu
�" .sysodisAaleR        TEXTt m   Q Tvv �ww 
 E r r o ru �!xy
�! 
mesSx l  W \z� �z b   W \{|{ m   W Z}} �~~ . F a i l e d   t o   c r e a t e   f o l d e r| o   Z [�� 0 yearrecvstr yearRecvStr�   �  y ��
� 
as A m   _ b�
� EAlTcriT�  s ��� L   i k�� m   i j��  �  b ��� r   l r��� n   l p��� 4   m p��
� 
cFld� o   n o�� 0 yearrecvstr yearRecvStr� o   l m�� 0 
destfolder 
destFolder� o      �� 0 	newfolder 	newFolder� ��� I  s z���
� .coremovenull���     obj � o   s t�� 0 
themessage 
theMessage� ���
� 
insh� o   u v�� 0 	newfolder 	newFolder�  �  �E  + m     ���                                                                                  OPIM  alis    D  gala-gs                        BD ����Microsoft Outlook.app                                          ����            ����  
 cu             Applications  %/:Applications:Microsoft Outlook.app/   ,  M i c r o s o f t   O u t l o o k . a p p    g a l a - g s  "Applications/Microsoft Outlook.app  / ��  # ��� l     ����  �  �  �       �
����
  � �	��	  0 archivemessage archiveMessage
� .aevtoappnull  �   � ****� �%������  0 archivemessage archiveMessage� ��� �  ��� 0 
themessage 
theMessage� 0 
destfolder 
destFolder�  � 	� �����������������  0 
themessage 
theMessage�� 0 
destfolder 
destFolder�� 0 daterecv dateRecv�� 0 yearrecv yearRecv�� 0 yearrecvstr yearRecvStr�� 0 	thefolder 	theFolder�� 0 errormessage errorMessage�� 0 errornumber errorNumber�� 0 	newfolder 	newFolder� �������G�����������������������v��}��������
�� 
pRed
�� 
rTim
�� 
year
�� 
cFld
�� .coredoexnull���     obj 
�� 
insh
�� .coremovenull���     obj 
�� 
kocl
�� 
cMFo
�� 
prdt
�� 
pnam�� 
�� .corecrel****      � null�� 0 errormessage errorMessage� ������
�� 
errn�� 0 errornumber errorNumber��  
�� 
mesS
�� 
as A
�� EAlTcriT�� 
�� .sysodisAaleR        TEXT� |� xe��,FO��,E�O��,E�O�%E�O��/j  ��/E�O��l Y D *�����l� W !X  a a a �%a a a  OjO��/E�O��l U� �����������
�� .aevtoappnull  �   � ****� k    z��  ��  ��  ��  ��  #����  ��  ��  � �������� 0 errormessage errorMessage�� 0 errornumber errorNumber�� 0 
themessage 
theMessage� 9 
�� �� �� �������������� F�� I�������������� s�� {����� � ��������������� � ����������������������� ��� 0 myname myName�� 0 mailname mailName�� 0 	inboxname 	inboxName�� 0 archivename archiveName
�� 
cwin�� 0 frontwin frontWin
�� 
pnam�� 0 winname winName
�� 
CMgs�� 0 currmsgs currMsgs
�� 
mesS
�� .sysodisAaleR        TEXT
�� 
cobj�� 0 firstmsg firstMsg
�� 
pOMC�� 0 onmycomputer onMyComputer
�� 
cFld�� 0 archivefolder archiveFolder��  0 destsentfolder destSentFolder��  0 destrecvfolder destRecvFolder�� 0 errormessage errorMessage� ������
�� 
errn�� 0 errornumber errorNumber��  
�� 
as A
�� EAlTcriT�� 
�� .corecnte****       ****�� 0 msgcount msgCount
�� 
appr
�� 
subt
�� .sysonotfnull��� ��� TEXT
�� 
dfAc�� 0 
defaccount 
defAccount
�� 
emad��  0 defsenderemail defSenderEmail
�� 
kocl
�� 
sndr�� 0 	senderobj 	senderObj
�� 
radd�� 0 senderemail senderEmail��  0 archivemessage archiveMessage��{�E�O�E�O�E�O�E�O�g*�k/E�O��,E�O*�,E�O�jv  ��%a a l OjY hO�a k/E` O*a ,E` O /_ a �/E` O_ a a /E` O_ a a /E` W #X  a a �a  %�%a !a "a # OjO�a -j $E` %O*a &�a 'a (_ %%a )%�%a # *O*a +,E` ,O_ ,a -,E` .O S M�[a /a l $kh �a 0,E` 1O_ 1a 2,E` 3O_ 3_ .  )�_ l+ 4Y )�_ l+ 4[OY��W #X  a 5a �a 6%�%a !a "a # OjO*a &�a 'a 7_ %%a 8%a # *Uascr  ��ޭ