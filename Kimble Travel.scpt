FasdUAS 1.101.10   ��   ��    k             l     ��  ��    8 2 Open Kimble in Safari (assumes you are logged in)     � 	 	 d   O p e n   K i m b l e   i n   S a f a r i   ( a s s u m e s   y o u   a r e   l o g g e d   i n )   
  
 l     ��  ��      How this works...     �   $   H o w   t h i s   w o r k s . . .      l     ��  ��    \ V Use the csv template to enter the details of the travel you wish to request in Kimble     �   �   U s e   t h e   c s v   t e m p l a t e   t o   e n t e r   t h e   d e t a i l s   o f   t h e   t r a v e l   y o u   w i s h   t o   r e q u e s t   i n   K i m b l e      l     ��  ��    o i Do NOT muck around with the headings, and don't use any commas. Make sure you save it as a csv, NOT xlsx     �   �   D o   N O T   m u c k   a r o u n d   w i t h   t h e   h e a d i n g s ,   a n d   d o n ' t   u s e   a n y   c o m m a s .   M a k e   s u r e   y o u   s a v e   i t   a s   a   c s v ,   N O T   x l s x      l     ��  ��    w q Save that file (don't rename it) in the same folder that this app will live in.  It expects it to be co-located.     �   �   S a v e   t h a t   f i l e   ( d o n ' t   r e n a m e   i t )   i n   t h e   s a m e   f o l d e r   t h a t   t h i s   a p p   w i l l   l i v e   i n .     I t   e x p e c t s   i t   t o   b e   c o - l o c a t e d .      l     ��������  ��  ��       !   l     "���� " r      # $ # l     %���� % I    �� &��
�� .earsffdralis        afdr &  f     ��  ��  ��   $ o      ���� "0 containerfolder containerFolder��  ��   !  ' ( ' l    )���� ) I   �� *��
�� .ascrcmnt****      � **** * o    	���� "0 containerfolder containerFolder��  ��  ��   (  + , + l    -���� - r     . / . c     0 1 0 l    2���� 2 b     3 4 3 l    5���� 5 n     6 7 6 1    ��
�� 
psxp 7 o    ���� "0 containerfolder containerFolder��  ��   4 m     8 8 � 9 9 . k i m b l e - t r a v e l - b a t c h . c s v��  ��   1 m    ��
�� 
psxf / o      ���� 0 	batchpath 	batchPath��  ��   ,  : ; : l     ��������  ��  ��   ;  < = < l    >���� > r     ? @ ? I   �� A��
�� .rdwrread****        **** A o    ���� 0 	batchpath 	batchPath��   @ o      ���� 0 csvtext csvText��  ��   =  B C B l    0 D���� D r     0 E F E I     .�� G���� 0 	csvtolist 	csvToList G  H I H o   ! "���� 0 csvtext csvText I  J K J K   " & L L �� M���� 0 	separator   M m   # $ N N � O O  ,��   K  P�� P K   & * Q Q �� R���� 0 trimming   R m   ' (��
�� boovtrue��  ��  ��   F o      ����  0 listofrequests listOfRequests��  ��   C  S T S l     ��������  ��  ��   T  U V U l  1 = W���� W r   1 = X Y X I  1 9�� Z��
�� .corecnte****       **** Z l  1 5 [���� [ n   1 5 \ ] \ 4   2 5�� ^
�� 
cobj ^ m   3 4����  ] o   1 2����  0 listofrequests listOfRequests��  ��  ��   Y o      ���� 0 numberofcols numberOfCols��  ��   V  _ ` _ l  > R a���� a Z  > R b c���� b >  > E d e d o   > A���� 0 numberofcols numberOfCols e m   A D���� 
 c R   H N�� f��
�� .ascrerr ****      � **** f m   J M g g � h h � O h   d e a r .   I t   l o o k s   l i k e   y o u ' v e   d o n e   s o m e t h i n g   i f f y   t o   t h e   c s v   f i l e ! !��  ��  ��  ��  ��   `  i j i l  S [ k���� k r   S [ l m l n   S W n o n 4   T W�� p
�� 
cobj p m   U V����  o o   S T����  0 listofrequests listOfRequests m o      ����  0 requestheaders requestHeaders��  ��   j  q r q l  \ k s���� s r   \ k t u t n   \ g v w v 7  ] g�� x y
�� 
cobj x m   a c����  y m   d f������ w o   \ ]����  0 listofrequests listOfRequests u o      ���� .0 listofrequestsnoheads listOfRequestsNoHeads��  ��   r  z { z l  l w |���� | r   l w } ~ } I  l s�� ��
�� .corecnte****       ****  o   l o���� .0 listofrequestsnoheads listOfRequestsNoHeads��   ~ o      ���� $0 numberofrequests numberOfRequests��  ��   {  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l  x  ����� � I  x �� ���
�� .ascrcmnt****      � **** � o   x {���� $0 numberofrequests numberOfRequests��  ��  ��   �  � � � l  � � ����� � I  � ��� ���
�� .ascrcmnt****      � **** � o   � ����� 0 numberofcols numberOfCols��  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l  � ����� � X   � ��� � � k   �z � �  � � � r   � � � � � n   � � � � � 4   � ��� �
�� 
cobj � m   � �����  � o   � ����� 0 
therequest 
theRequest � o      ���� 0 maintraveller mainTraveller �  � � � r   � � � � � n   � � � � � 4   � ��� �
�� 
cobj � m   � �����  � o   � ����� 0 
therequest 
theRequest � o      ���� 0 travelsummary travelSummary �  � � � r   � � � � � n   � � � � � 4   � ��� �
�� 
cobj � m   � �����  � o   � ����� 0 
therequest 
theRequest � o      �� 0 fromdate fromDate �  � � � r   � � � � � n   � � � � � 4   � ��~ �
�~ 
cobj � m   � ��}�}  � o   � ��|�| 0 
therequest 
theRequest � o      �{�{ 0 todate toDate �  � � � r   � � � � � n   � � � � � 4   � ��z �
�z 
cobj � m   � ��y�y  � o   � ��x�x 0 
therequest 
theRequest � o      �w�w 0 activity   �  � � � r   � � � � � n   � � � � � 4   � ��v �
�v 
cobj � m   � ��u�u  � o   � ��t�t 0 
therequest 
theRequest � o      �s�s 
0 reason   �  � � � r   � � � � � c   � � � � � n   � � � � � 4   � ��r �
�r 
cobj � m   � ��q�q  � o   � ��p�p 0 
therequest 
theRequest � m   � ��o
�o 
bool � o      �n�n 0 flightbaggage flightBaggage �  � � � r   � � � � � c   � � � � � n   � � � � � 4   � ��m �
�m 
cobj � m   � ��l�l  � o   � ��k�k 0 
therequest 
theRequest � m   � ��j
�j 
bool � o      �i�i 0 
flightneed 
flightNeed �  � � � r   � � � � c   �  � � � n   � � � � � 4   � ��h �
�h 
cobj � m   � ��g�g 	 � o   � ��f�f 0 
therequest 
theRequest � m   � ��e
�e 
bool � o      �d�d 0 	hotelneed 	hotelNeed �  � � � r   � � � c   � � � n   � � � 4  �c �
�c 
cobj � m  
�b�b 
 � o  �a�a 0 
therequest 
theRequest � m  �`
�` 
bool � o      �_�_ 0 	trainneed 	trainNeed �  � � � l �^�]�\�^  �]  �\   �  � � � O  t � � � k  s � �  � � � O  B � � � r  #A � � � l #; ��[�Z � I #;�Y�X �
�Y .corecrel****      � null�X   � �W � �
�W 
kocl � m  '*�V
�V 
bTab � �U ��T
�U 
prdt � K  -5 � � �S ��R
�S 
pURL � m  03 � � � � � @ h t t p s : / / e u 1 . s a l e s f o r c e . c o m / a 3 2 / o�R  �T  �[  �Z   � 1  ;@�Q
�Q 
cTab � 4   �P �
�P 
cwin � m  �O�O  �  � � � I CH�N ��M
�N .sysodelanull��� ��� nmbr � m  CD�L�L �M   �  � � � l II�K �K    . (	DatePicker.datePicker.selectDate(this);    � P 	 D a t e P i c k e r . d a t e P i c k e r . s e l e c t D a t e ( t h i s ) ; �  I IY�J
�J .sfridojs****       utxt m  IL � P d o c u m e n t . f o r m s [ ' h o t l i s t ' ] [ ' n e w ' ] . c l i c k ( ) �I	�H
�I 
dcnm	 4  OU�G

�G 
docu
 m  ST�F�F �H    I Z_�E�D
�E .sysodelanull��� ��� nmbr m  Z[�C�C �D    I `x�B
�B .sfridojs****       utxt b  `k b  `g m  `c � � d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 0 9 : j _ i d 1 1 2 ' ] . v a l u e   =   ' o  cf�A�A 0 maintraveller mainTraveller m  gj �  ' �@�?
�@ 
dcnm 4  nt�>
�> 
docu m  rs�=�= �?    l yy�<�<   n h	do JavaScript "document.forms['j_id0:j_id1:TheForm'][''].value = '" & mainTraveller & "'" in document 1    �   � 	 d o   J a v a S c r i p t   " d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' ' ] . v a l u e   =   ' "   &   m a i n T r a v e l l e r   &   " ' "   i n   d o c u m e n t   1 !"! I y��;#$
�; .sfridojs****       utxt# b  y�%&% b  y�'(' m  y|)) �** � d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 1 7 : j _ i d 1 1 9 ' ] . v a l u e   =   '( o  |�:�: 0 travelsummary travelSummary& m  ��++ �,,  '$ �9-�8
�9 
dcnm- 4  ���7.
�7 
docu. m  ���6�6 �8  " /0/ l ���5�4�3�5  �4  �3  0 121 I ���234
�2 .sfridojs****       utxt3 b  ��565 b  ��787 m  ��99 �:: . D a t e P i c k e r . i n s e r t D a t e ( '8 o  ���1�1 0 fromdate fromDate6 m  ��;; �<< ~ ' ,   ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 2 0 : j _ i d 1 2 3 ' , t r u e ) ;4 �0=�/
�0 
dcnm= 4  ���.>
�. 
docu> m  ���-�- �/  2 ?@? I ���,AB
�, .sfridojs****       utxtA b  ��CDC b  ��EFE m  ��GG �HH . D a t e P i c k e r . i n s e r t D a t e ( 'F o  ���+�+ 0 todate toDateD m  ��II �JJ � ' ,   ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 2 5 : j _ i d 1 2 8 ' , f a l s e ) ;B �*K�)
�* 
dcnmK 4  ���(L
�( 
docuL m  ���'�' �)  @ MNM l ���&�%�$�&  �%  �$  N OPO l ���#QR�#  Q � �	do JavaScript "document.forms['j_id0:j_id1:TheForm']['j_id0:j_id1:TheForm:j_id104:j_id108:j_id125:j_id128'].value = '" & toDate & "'" in document 1   R �SS( 	 d o   J a v a S c r i p t   " d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 2 5 : j _ i d 1 2 8 ' ] . v a l u e   =   ' "   &   t o D a t e   &   " ' "   i n   d o c u m e n t   1P TUT I ���"V�!
�" .sysodelanull��� ��� nmbrV m  ��� �  �!  U WXW I ���YZ
� .sfridojs****       utxtY m  ��[[ �\\ � d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 3 0 : A c t i v i t y L i s t ' ] . s e l e c t e d I n d e x   =   ' 5 ' ;Z �]�
� 
dcnm] 4  ���^
� 
docu^ m  ���� �  X _`_ I ���ab
� .sfridojs****       utxta b  ��cdc b  ��efe m  ��gg �hh � d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 0 8 : j _ i d 1 3 8 : j _ i d 1 4 0 ' ] . v a l u e   =   'f o  ���� 
0 reason  d m  ��ii �jj  ' ;b �k�
� 
dcnmk 4  ���l
� 
docul m  ���� �  ` mnm Z  �op��o = ��qrq o  ���� 0 flightbaggage flightBaggager m  ���
� boovtruep I ��st
� .sfridojs****       utxts m  ��uu �vv � d o c u m e n t . f o r m s [ ' j _ i d 0 : j _ i d 1 : T h e F o r m ' ] [ ' j _ i d 0 : j _ i d 1 : T h e F o r m : j _ i d 1 0 4 : j _ i d 1 4 1 : j _ i d 1 4 2 : 0 : j _ i d 1 4 3 ' ] . c h e c k e d   =   t r u e ;t �w�
� 
dcnmw 4  �x
� 
docux m  �� �  �  �  n yzy l �{|�  { ] W addRequisitionItem('Flight')  addRequisitionItem('Hotel')  addRequisitionItem('Train')   | �}} �   a d d R e q u i s i t i o n I t e m ( ' F l i g h t ' )     a d d R e q u i s i t i o n I t e m ( ' H o t e l ' )     a d d R e q u i s i t i o n I t e m ( ' T r a i n ' )z ~~ Z  1���
�	� o  �� 0 
flightneed 
flightNeed� k  -�� ��� I '���
� .sfridojs****       utxt� m  �� ��� 8 a d d R e q u i s i t i o n I t e m ( ' F l i g h t ' )� ���
� 
dcnm� 4  #��
� 
docu� m  !"�� �  � ��� I (-��� 
� .sysodelanull��� ��� nmbr� m  ()���� �   �  �
  �	   ��� Z  2R������� o  25���� 0 	hotelneed 	hotelNeed� k  8N�� ��� I 8H����
�� .sfridojs****       utxt� m  8;�� ��� 6 a d d R e q u i s i t i o n I t e m ( ' H o t e l ' )� �����
�� 
dcnm� 4  >D���
�� 
docu� m  BC���� ��  � ���� I IN�����
�� .sysodelanull��� ��� nmbr� m  IJ���� ��  ��  ��  ��  � ���� Z  Ss������� o  SV���� 0 	trainneed 	trainNeed� k  Yo�� ��� I Yi����
�� .sfridojs****       utxt� m  Y\�� ��� 6 a d d R e q u i s i t i o n I t e m ( ' T r a i n ' )� �����
�� 
dcnm� 4  _e���
�� 
docu� m  cd���� ��  � ���� I jo�����
�� .sysodelanull��� ��� nmbr� m  jk���� ��  ��  ��  ��  ��   � m  ���                                                                                  sfri  alis    N  Macintosh HD               �S/�H+   _�
Safari.app                                                      `ERѮ-�        ����  	                Applications    �S!�      Ѯ�     _�  %Macintosh HD:Applications: Safari.app    
 S a f a r i . a p p    M a c i n t o s h   H D  Applications/Safari.app   / ��   � ���� I uz�����
�� .sysodelanull��� ��� nmbr� m  uv���� ��  ��  �� 0 
therequest 
theRequest � o   � ����� .0 listofrequestsnoheads listOfRequestsNoHeads��  ��   � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � h b Following cvv to list-of-list script from http://macscripter.net/viewtopic.php?pid=125444#p125444   � ��� �   F o l l o w i n g   c v v   t o   l i s t - o f - l i s t   s c r i p t   f r o m   h t t p : / / m a c s c r i p t e r . n e t / v i e w t o p i c . p h p ? p i d = 1 2 5 4 4 4 # p 1 2 5 4 4 4� ��� l     ��������  ��  ��  � ��� l      ������  � Assumes that the CSV text adheres to the convention:
   Records are delimited by LFs or CRLFs (but CRs are also allowed here).
   The last record in the text may or may not be followed by an LF or CRLF (or CR).
   Fields in the same record are separated by commas (unless specified differently by parameter).
   The last field in a record must not be followed by a comma.
   Trailing or leading spaces in unquoted fields are not ignored (unless so specified by parameter).
   Fields containing quoted text are quoted in their entirety, any space outside them being ignored.
   Fields enclosed in double-quotes are to be taken verbatim, except for any included double-quote pairs, which are to be translated as double-quote characters.
       
   No other variations are currently supported.    � ���0   A s s u m e s   t h a t   t h e   C S V   t e x t   a d h e r e s   t o   t h e   c o n v e n t i o n : 
       R e c o r d s   a r e   d e l i m i t e d   b y   L F s   o r   C R L F s   ( b u t   C R s   a r e   a l s o   a l l o w e d   h e r e ) . 
       T h e   l a s t   r e c o r d   i n   t h e   t e x t   m a y   o r   m a y   n o t   b e   f o l l o w e d   b y   a n   L F   o r   C R L F   ( o r   C R ) . 
       F i e l d s   i n   t h e   s a m e   r e c o r d   a r e   s e p a r a t e d   b y   c o m m a s   ( u n l e s s   s p e c i f i e d   d i f f e r e n t l y   b y   p a r a m e t e r ) . 
       T h e   l a s t   f i e l d   i n   a   r e c o r d   m u s t   n o t   b e   f o l l o w e d   b y   a   c o m m a . 
       T r a i l i n g   o r   l e a d i n g   s p a c e s   i n   u n q u o t e d   f i e l d s   a r e   n o t   i g n o r e d   ( u n l e s s   s o   s p e c i f i e d   b y   p a r a m e t e r ) . 
       F i e l d s   c o n t a i n i n g   q u o t e d   t e x t   a r e   q u o t e d   i n   t h e i r   e n t i r e t y ,   a n y   s p a c e   o u t s i d e   t h e m   b e i n g   i g n o r e d . 
       F i e l d s   e n c l o s e d   i n   d o u b l e - q u o t e s   a r e   t o   b e   t a k e n   v e r b a t i m ,   e x c e p t   f o r   a n y   i n c l u d e d   d o u b l e - q u o t e   p a i r s ,   w h i c h   a r e   t o   b e   t r a n s l a t e d   a s   d o u b l e - q u o t e   c h a r a c t e r s . 
               
       N o   o t h e r   v a r i a t i o n s   a r e   c u r r e n t l y   s u p p o r t e d .  � ��� l     ��������  ��  ��  � ��� i     ��� I      ������� 0 	csvtolist 	csvToList� ��� o      ���� 0 csvtext csvText� ���� o      ���� 0 implementation  ��  ��  � k    <�� ��� l     ������  �YS The 'implementation' parameter must be a record. Leave it empty ({}) for the default assumptions: ie. comma separator, leading and trailing spaces in unquoted fields not to be trimmed. Otherwise it can have a 'separator' property with a text value (eg. {separator:tab}) and/or a 'trimming' property with a boolean value ({trimming:true}).   � ����   T h e   ' i m p l e m e n t a t i o n '   p a r a m e t e r   m u s t   b e   a   r e c o r d .   L e a v e   i t   e m p t y   ( { } )   f o r   t h e   d e f a u l t   a s s u m p t i o n s :   i e .   c o m m a   s e p a r a t o r ,   l e a d i n g   a n d   t r a i l i n g   s p a c e s   i n   u n q u o t e d   f i e l d s   n o t   t o   b e   t r i m m e d .   O t h e r w i s e   i t   c a n   h a v e   a   ' s e p a r a t o r '   p r o p e r t y   w i t h   a   t e x t   v a l u e   ( e g .   { s e p a r a t o r : t a b } )   a n d / o r   a   ' t r i m m i n g '   p r o p e r t y   w i t h   a   b o o l e a n   v a l u e   ( { t r i m m i n g : t r u e } ) .� ��� r     ��� l    ������ b     ��� o     ���� 0 implementation  � K    �� ������ 0 	separator  � m    �� ���  ,� ������� 0 trimming  � m    ��
�� boovfals��  ��  ��  � K      �� ������ 0 	separator  � o      ���� 0 	separator  � ������� 0 trimming  � o      ���� 0 trimming  ��  � ��� l   ��������  ��  ��  � ��� h    ����� 0 o  � l     ���� k      �� ��� j     ����� 0 qdti  � I     ������� 0 gettextitems getTextItems� ��� o    ���� 0 csvtext csvText� ���� m    �� ���  "��  ��  � ��� j    ����� 0 currentrecord currentRecord� J    ����  � ��� j    �����  0 possiblefields possibleFields� m    ��
�� 
msng� ���� j    ����� 0 
recordlist 
recordList� J    ����  ��  �   Lists for fast access.   � ��� .   L i s t s   f o r   f a s t   a c c e s s .� � � l   ��������  ��  ��     l   ����   Q K o's qdti is a list of the CSV's text items, as delimited by double-quotes.    � �   o ' s   q d t i   i s   a   l i s t   o f   t h e   C S V ' s   t e x t   i t e m s ,   a s   d e l i m i t e d   b y   d o u b l e - q u o t e s .  l   ��	��   R L Assuming the convention mentioned above, the number of items is always odd.   	 �

 �   A s s u m i n g   t h e   c o n v e n t i o n   m e n t i o n e d   a b o v e ,   t h e   n u m b e r   o f   i t e m s   i s   a l w a y s   o d d .  l   ����   S M Even-numbered items (if any) are quoted field values and don't need parsing.    � �   E v e n - n u m b e r e d   i t e m s   ( i f   a n y )   a r e   q u o t e d   f i e l d   v a l u e s   a n d   d o n ' t   n e e d   p a r s i n g .  l   ����   R L Odd-numbered items are everything else. Empty strings in odd-numbered slots    � �   O d d - n u m b e r e d   i t e m s   a r e   e v e r y t h i n g   e l s e .   E m p t y   s t r i n g s   i n   o d d - n u m b e r e d   s l o t s  l   ����   R L (except at the beginning and end) indicate escaped quotes in quoted fields.    � �   ( e x c e p t   a t   t h e   b e g i n n i n g   a n d   e n d )   i n d i c a t e   e s c a p e d   q u o t e s   i n   q u o t e d   f i e l d s .  l   ��������  ��  ��    r    # n   ! !  1    !��
�� 
txdl! 1    ��
�� 
ascr o      ���� 	0 astid   "#" r   $ -$%$ l  $ +&����& I  $ +��'��
�� .corecnte****       ****' n  $ '()( o   % '���� 0 qdti  ) o   $ %���� 0 o  ��  ��  ��  % o      ���� 0 	qdticount 	qdtiCount# *+* r   . 1,-, m   . /��
�� boovfals- o      ���� "0 quoteinprogress quoteInProgress+ ./. P   201��0 Y   72��3452 l  A6786 k   A99 :;: r   A I<=< n   A G>?> 4   D G��@
�� 
cobj@ o   E F���� 0 i  ? n  A DABA o   B D���� 0 qdti  B o   A B���� 0 o  = o      ���� 0 thisbit thisBit; C��C Z   JDEF��D G   J YGHG l  J QI����I ?   J QJKJ l  J OL����L I  J O��M��
�� .corecnte****       ****M o   J K���� 0 thisbit thisBit��  ��  ��  K m   O P����  ��  ��  H l  T WN��~N =  T WOPO o   T U�}�} 0 i  P o   U V�|�| 0 	qdticount 	qdtiCount�  �~  E k   \�QQ RSR l  \ \�{TU�{  T T N This is either a non-empty string or the last item in the list, so it doesn't   U �VV �   T h i s   i s   e i t h e r   a   n o n - e m p t y   s t r i n g   o r   t h e   l a s t   i t e m   i n   t h e   l i s t ,   s o   i t   d o e s n ' tS WXW l  \ \�zYZ�z  Y K E represent a quoted quote. Check if we've just been dealing with any.   Z �[[ �   r e p r e s e n t   a   q u o t e d   q u o t e .   C h e c k   i f   w e ' v e   j u s t   b e e n   d e a l i n g   w i t h   a n y .X \]\ Z   \ �^_`�y^ l  \ ]a�x�wa o   \ ]�v�v "0 quoteinprogress quoteInProgress�x  �w  _ k   ` �bb cdc l  ` `�uef�u  e M G All the parts of a quoted field containing quoted quotes have now been   f �gg �   A l l   t h e   p a r t s   o f   a   q u o t e d   f i e l d   c o n t a i n i n g   q u o t e d   q u o t e s   h a v e   n o w   b e e nd hih l  ` `�tjk�t  j A ; passed over. Coerce them together using a quote delimiter.   k �ll v   p a s s e d   o v e r .   C o e r c e   t h e m   t o g e t h e r   u s i n g   a   q u o t e   d e l i m i t e r .i mnm r   ` eopo m   ` aqq �rr  "p n     sts 1   b d�s
�s 
txdlt 1   a b�r
�r 
ascrn uvu r   f ywxw c   f wyzy l  f u{�q�p{ n   f u|}| 7  i u�o~
�o 
cobj~ o   m o�n�n 0 a   l  p t��m�l� \   p t��� o   q r�k�k 0 i  � m   r s�j�j �m  �l  } n  f i��� o   g i�i�i 0 qdti  � o   f g�h�h 0 o  �q  �p  z m   u v�g
�g 
TEXTx o      �f�f 0 	thisfield 	thisFieldv ��� l  z z�e���e  � C = Replace the reconstituted quoted quotes with literal quotes.   � ��� z   R e p l a c e   t h e   r e c o n s t i t u t e d   q u o t e d   q u o t e s   w i t h   l i t e r a l   q u o t e s .� ��� r   z ���� m   z }�� ���  " "� n     ��� 1   ~ ��d
�d 
txdl� 1   } ~�c
�c 
ascr� ��� r   � ���� n  � ���� 2  � ��b
�b 
citm� o   � ��a�a 0 	thisfield 	thisField� o      �`�` 0 	thisfield 	thisField� ��� r   � ���� m   � ��� ���  "� n     ��� 1   � ��_
�_ 
txdl� 1   � ��^
�^ 
ascr� ��� l  � ��]���]  � \ V Store the field in the "current record" list and cancel the "quote in progress" flag.   � ��� �   S t o r e   t h e   f i e l d   i n   t h e   " c u r r e n t   r e c o r d "   l i s t   a n d   c a n c e l   t h e   " q u o t e   i n   p r o g r e s s "   f l a g .� ��� r   � ���� c   � ���� o   � ��\�\ 0 	thisfield 	thisField� m   � ��[
�[ 
TEXT� n      ���  ;   � �� n  � ���� o   � ��Z�Z 0 currentrecord currentRecord� o   � ��Y�Y 0 o  � ��X� r   � ���� m   � ��W
�W boovfals� o      �V�V "0 quoteinprogress quoteInProgress�X  ` ��� l  � ���U�T� ?   � ���� o   � ��S�S 0 i  � m   � ��R�R �U  �T  � ��Q� k   � ��� ��� l  � ��P���P  � N H The preceding, even-numbered item is a complete quoted field. Store it.   � ��� �   T h e   p r e c e d i n g ,   e v e n - n u m b e r e d   i t e m   i s   a   c o m p l e t e   q u o t e d   f i e l d .   S t o r e   i t .� ��O� r   � ���� n   � ���� 4   � ��N�
�N 
cobj� l  � ���M�L� \   � ���� o   � ��K�K 0 i  � m   � ��J�J �M  �L  � n  � ���� o   � ��I�I 0 qdti  � o   � ��H�H 0 o  � n      ���  ;   � �� n  � ���� o   � ��G�G 0 currentrecord currentRecord� o   � ��F�F 0 o  �O  �Q  �y  ] ��� l  � ��E�D�C�E  �D  �C  � ��� l  � ��B���B  �60 Now parse this item's field-separator-delimited text items, which are either non-quoted fields or stumps from the removal of quoted fields. Any that contain line breaks must be further split to end one record and start another. These could include multiple single-field records without field separators.   � ���`   N o w   p a r s e   t h i s   i t e m ' s   f i e l d - s e p a r a t o r - d e l i m i t e d   t e x t   i t e m s ,   w h i c h   a r e   e i t h e r   n o n - q u o t e d   f i e l d s   o r   s t u m p s   f r o m   t h e   r e m o v a l   o f   q u o t e d   f i e l d s .   A n y   t h a t   c o n t a i n   l i n e   b r e a k s   m u s t   b e   f u r t h e r   s p l i t   t o   e n d   o n e   r e c o r d   a n d   s t a r t   a n o t h e r .   T h e s e   c o u l d   i n c l u d e   m u l t i p l e   s i n g l e - f i e l d   r e c o r d s   w i t h o u t   f i e l d   s e p a r a t o r s .� ��� r   � ���� I   � ��A��@�A 0 gettextitems getTextItems� ��� o   � ��?�? 0 thisbit thisBit� ��>� o   � ��=�= 0 	separator  �>  �@  � n     ��� o   � ��<�<  0 possiblefields possibleFields� o   � ��;�; 0 o  � ��� r   � ���� l  � ���:�9� I  � ��8��7
�8 .corecnte****       ****� n  � ���� o   � ��6�6  0 possiblefields possibleFields� o   � ��5�5 0 o  �7  �:  �9  � o      �4�4 (0 possiblefieldcount possibleFieldCount� ��� Y   ����3���2� k   ���� ��� r   � ���� n   � ���� 4   � ��1�
�1 
cobj� o   � ��0�0 0 j  � n  � ���� o   � ��/�/  0 possiblefields possibleFields� o   � ��.�. 0 o  � o      �-�- 0 	thisfield 	thisField� ��,� Z   �����+�� l  � ���*�)� ?   � ���� l  � ���(�'� I  � ��&��
�& .corecnte****       ****� o   � ��%�% 0 	thisfield 	thisField� �$��#
�$ 
kocl� m   � ��"
�" 
cpar�#  �(  �'  � m   � ��!�! �*  �)  � k   ��    l  � �� �    P J This "field" contains one or more line endings. Split it at those points.    � �   T h i s   " f i e l d "   c o n t a i n s   o n e   o r   m o r e   l i n e   e n d i n g s .   S p l i t   i t   a t   t h o s e   p o i n t s .  r   �	 n  �

 2  ��
� 
cpar o   � ��� 0 	thisfield 	thisField	 o      �� 0 thesefields theseFields  l ��   � � With each of these end-of-record fields except the last, complete the field list for the current record and initialise another. Omit the first "field" if it's just the stub from a preceding quoted field.    ��   W i t h   e a c h   o f   t h e s e   e n d - o f - r e c o r d   f i e l d s   e x c e p t   t h e   l a s t ,   c o m p l e t e   t h e   f i e l d   l i s t   f o r   t h e   c u r r e n t   r e c o r d   a n d   i n i t i a l i s e   a n o t h e r .   O m i t   t h e   f i r s t   " f i e l d "   i f   i t ' s   j u s t   t h e   s t u b   f r o m   a   p r e c e d i n g   q u o t e d   f i e l d .  Y  p�� k  k  r   n   4  �
� 
cobj o  �� 0 k   o  �� 0 thesefields theseFields o      �� 0 	thisfield 	thisField  Z U !��  l @"��" G  @#$# G  .%&% G  &'(' l )��) ?  *+* o  �� 0 k  + m  �� �  �  ( l !$,��, ?  !$-.- o  !"�� 0 j  . m  "#�
�
 �  �  & l ),/�	�/ = ),010 o  )*�� 0 i  1 m  *+�� �	  �  $ l 1>2��2 ?  1>343 l 1<5��5 I 1<�6� 
� .corecnte****       ****6 I  18��7���� 0 trim  7 898 o  23���� 0 	thisfield 	thisField9 :��: m  34��
�� boovtrue��  ��  �   �  �  4 m  <=����  �  �  �  �  ! r  CQ;<; I  CJ��=���� 0 trim  = >?> o  DE���� 0 	thisfield 	thisField? @��@ o  EF���� 0 trimming  ��  ��  < n      ABA  ;  OPB n JOCDC o  KO���� 0 currentrecord currentRecordD o  JK���� 0 o  �  �   EFE r  VbGHG n V[IJI o  W[���� 0 currentrecord currentRecordJ o  VW���� 0 o  H n      KLK  ;  `aL n [`MNM o  \`���� 0 
recordlist 
recordListN o  [\���� 0 o  F O��O r  ckPQP J  ce����  Q n     RSR o  fj���� 0 currentrecord currentRecordS o  ef���� 0 o  ��  � 0 k   m  ����  \  TUT l V����V I ��W��
�� .corecnte****       ****W o  	���� 0 thesefields theseFields��  ��  ��  U m  ���� �   XYX l qq��Z[��  Z � � With the last end-of-record "field", just complete the current field list if the field's not the stub from a following quoted field.   [ �\\
   W i t h   t h e   l a s t   e n d - o f - r e c o r d   " f i e l d " ,   j u s t   c o m p l e t e   t h e   c u r r e n t   f i e l d   l i s t   i f   t h e   f i e l d ' s   n o t   t h e   s t u b   f r o m   a   f o l l o w i n g   q u o t e d   f i e l d .Y ]^] r  qu_`_ n  qsaba  ;  rsb o  qr���� 0 thesefields theseFields` o      ���� 0 	thisfield 	thisField^ c��c Z v�de����d l v�f����f G  v�ghg l vyi����i A  vyjkj o  vw���� 0 j  k o  wx���� (0 possiblefieldcount possibleFieldCount��  ��  h l |�l����l ?  |�mnm l |�o����o I |���p��
�� .corecnte****       ****p o  |}���� 0 	thisfield 	thisField��  ��  ��  n m  ������  ��  ��  ��  ��  e r  ��qrq I  ����s���� 0 trim  s tut o  ������ 0 	thisfield 	thisFieldu v��v o  ������ 0 trimming  ��  ��  r n      wxw  ;  ��x n ��yzy o  ������ 0 currentrecord currentRecordz o  ������ 0 o  ��  ��  ��  �+  � k  ��{{ |}| l ����~��  ~ � � This is a "field" not containing a line break. Insert it into the current field list if it's not just a stub from a preceding or following quoted field.    ���2   T h i s   i s   a   " f i e l d "   n o t   c o n t a i n i n g   a   l i n e   b r e a k .   I n s e r t   i t   i n t o   t h e   c u r r e n t   f i e l d   l i s t   i f   i t ' s   n o t   j u s t   a   s t u b   f r o m   a   p r e c e d i n g   o r   f o l l o w i n g   q u o t e d   f i e l d .} ���� Z ��������� l �������� G  ����� G  ����� l �������� F  ����� l �������� ?  ����� o  ������ 0 j  � m  ������ ��  ��  � l �������� G  ����� l �������� A  ����� o  ������ 0 j  � o  ������ (0 possiblefieldcount possibleFieldCount��  ��  � l �������� = ����� o  ������ 0 i  � o  ������ 0 	qdticount 	qdtiCount��  ��  ��  ��  ��  ��  � l �������� F  ����� l �������� = ����� o  ������ 0 j  � m  ������ ��  ��  � l �������� = ����� o  ������ 0 i  � m  ������ ��  ��  ��  ��  � l �������� ?  ����� l �������� I �������
�� .corecnte****       ****� I  ��������� 0 trim  � ��� o  ������ 0 	thisfield 	thisField� ���� m  ����
�� boovtrue��  ��  ��  ��  ��  � m  ������  ��  ��  ��  ��  � r  ����� I  ��������� 0 trim  � ��� o  ������ 0 	thisfield 	thisField� ���� o  ������ 0 trimming  ��  ��  � n      ���  ;  ��� n ����� o  ������ 0 currentrecord currentRecord� o  ������ 0 o  ��  ��  ��  �,  �3 0 j  � m   � ����� � o   � ����� (0 possiblefieldcount possibleFieldCount�2  � ��� l ����������  ��  ��  � ���� l ��������  � I C Otherwise, this item IS an empty text representing a quoted quote.   � ��� �   O t h e r w i s e ,   t h i s   i t e m   I S   a n   e m p t y   t e x t   r e p r e s e n t i n g   a   q u o t e d   q u o t e .��  F ��� l �������� o  ������ "0 quoteinprogress quoteInProgress��  ��  � ��� l ��������  � Z T It's another quote in a field already identified as having one. Do nothing for now.   � ��� �   I t ' s   a n o t h e r   q u o t e   i n   a   f i e l d   a l r e a d y   i d e n t i f i e d   a s   h a v i n g   o n e .   D o   n o t h i n g   f o r   n o w .� ��� l �������� ?  ����� o  ������ 0 i  � m  ������ ��  ��  � ���� k  ��� ��� l ��������  � K E It's the first quoted quote in a quoted field. Note the index of the   � ��� �   I t ' s   t h e   f i r s t   q u o t e d   q u o t e   i n   a   q u o t e d   f i e l d .   N o t e   t h e   i n d e x   o f   t h e� ��� l ��������  � T N preceding even-numbered item (the first part of the field) and flag "quote in   � ��� �   p r e c e d i n g   e v e n - n u m b e r e d   i t e m   ( t h e   f i r s t   p a r t   o f   t h e   f i e l d )   a n d   f l a g   " q u o t e   i n� ��� l ��������  � R L progress" so that the repeat idles past the remaining part(s) of the field.   � ��� �   p r o g r e s s "   s o   t h a t   t h e   r e p e a t   i d l e s   p a s t   t h e   r e m a i n i n g   p a r t ( s )   o f   t h e   f i e l d .� ��� r  ���� \  ���� o  � ���� 0 i  � m   ���� � o      ���� 0 a  � ��� r  ��� m  �~
�~ boovtrue� o      �}�} "0 quoteinprogress quoteInProgress�  ��  ��  ��  7 %  Parse odd-numbered items only.   8 ��� >   P a r s e   o d d - n u m b e r e d   i t e m s   o n l y .�� 0 i  3 m   : ;�|�| 4 o   ; <�{�{ 0 	qdticount 	qdtiCount5 m   < =�z�z 1 �y�x
�y conscase�x  ��  / ��� l �w�v�u�w  �v  �u  � ��� l �t���t  � F @ At the end of the repeat, store any remaining "current record".   � ��� �   A t   t h e   e n d   o f   t h e   r e p e a t ,   s t o r e   a n y   r e m a i n i n g   " c u r r e n t   r e c o r d " .� ��� Z .���s�r� l ��q�p� > ��� n ��� o  �o�o 0 currentrecord currentRecord� o  �n�n 0 o  � J  �m�m  �q  �p  � r  *��� n #��� o  #�l�l 0 currentrecord currentRecord� o  �k�k 0 o  � n      ���  ;  ()� n #(��� o  $(�j�j 0 
recordlist 
recordList� o  #$�i�i 0 o  �s  �r  � ��� r  /4��� o  /0�h�h 	0 astid  � n     ��� 1  13�g
�g 
txdl� 1  01�f
�f 
ascr�    l 55�e�d�c�e  �d  �c   �b L  5< n 5; o  6:�a�a 0 
recordlist 
recordList o  56�`�` 0 o  �b  �  l     �_�^�]�_  �^  �]   	 l     �\
�\  
 > 8 Get the possibly more than 4000 text items from a text.    � p   G e t   t h e   p o s s i b l y   m o r e   t h a n   4 0 0 0   t e x t   i t e m s   f r o m   a   t e x t .	  i     I      �[�Z�[ 0 gettextitems getTextItems  o      �Y�Y 0 txt   �X o      �W�W 	0 delim  �X  �Z   k     V  r      n     1    �V
�V 
txdl 1     �U
�U 
ascr o      �T�T 	0 astid    r     o    �S�S 	0 delim   n      !  1    
�R
�R 
txdl! 1    �Q
�Q 
ascr "#" r    $%$ l   &�P�O& I   �N'�M
�N .corecnte****       ****' n   ()( 2   �L
�L 
citm) o    �K�K 0 txt  �M  �P  �O  % o      �J�J 0 ticount tiCount# *+* r    ,-, J    �I�I  - o      �H�H 0 	textitems 	textItems+ ./. Y    M0�G1230 k   % H44 565 r   % *787 [   % (9:9 o   % &�F�F 0 i  : m   & '�E�E�8 o      �D�D 0 j  6 ;<; Z  + 8=>�C�B= l  + .?�A�@? ?   + .@A@ o   + ,�?�? 0 j  A o   , -�>�> 0 ticount tiCount�A  �@  > r   1 4BCB o   1 2�=�= 0 ticount tiCountC o      �<�< 0 j  �C  �B  < D�;D r   9 HEFE b   9 FGHG o   9 :�:�: 0 	textitems 	textItemsH n   : EIJI 7  ; E�9KL
�9 
citmK o   ? A�8�8 0 i  L o   B D�7�7 0 j  J o   : ;�6�6 0 txt  F o      �5�5 0 	textitems 	textItems�;  �G 0 i  1 m    �4�4 2 o     �3�3 0 ticount tiCount3 m     !�2�2�/ MNM r   N SOPO o   N O�1�1 	0 astid  P n     QRQ 1   P R�0
�0 
txdlR 1   O P�/
�/ 
ascrN STS l  T T�.�-�,�.  �-  �,  T U�+U L   T VVV o   T U�*�* 0 	textitems 	textItems�+   WXW l     �)�(�'�)  �(  �'  X YZY l     �&[\�&  [ 9 3 Trim any leading or trailing spaces from a string.   \ �]] f   T r i m   a n y   l e a d i n g   o r   t r a i l i n g   s p a c e s   f r o m   a   s t r i n g .Z ^�%^ i    _`_ I      �$a�#�$ 0 trim  a bcb o      �"�" 0 txt  c d�!d o      � �  0 trimming  �!  �#  ` k     ree fgf Z     ohi��h l    j��j o     �� 0 trimming  �  �  i k    kkk lml Y    0n�op�n Z    +qr�sq l   t��t C   uvu o    �� 0 txt  v 1    �
� 
spac�  �  r r    'wxw n    %yzy 7   %�{|
� 
ctxt{ m    !�� | m   " $����z o    �� 0 txt  x o      �� 0 txt  �  s  S   * +� 0 i  o m    �� p \    }~} l   �� I   ���

� .corecnte****       ****� o    	�	�	 0 txt  �
  �  �  ~ m    �� �  m ��� Y   1 ]������ Z   A X����� l  A D���� D   A D��� o   A B�� 0 txt  � 1   B C�
� 
spac�  �  � r   G T��� n   G R��� 7  H R� ��
�  
ctxt� m   L N���� � m   O Q������� o   G H���� 0 txt  � o      ���� 0 txt  �  �  S   W X� 0 i  � m   4 5���� � \   5 <��� l  5 :������ I  5 :�����
�� .corecnte****       ****� o   5 6���� 0 txt  ��  ��  ��  � m   : ;���� �  � ���� Z  ^ k������� l  ^ a������ =  ^ a��� o   ^ _���� 0 txt  � 1   _ `��
�� 
spac��  ��  � r   d g��� m   d e�� ���  � o      ���� 0 txt  ��  ��  ��  �  �  g ��� l  p p��������  ��  ��  � ���� L   p r�� o   p q���� 0 txt  ��  �%       ���������  � ���������� 0 	csvtolist 	csvToList�� 0 gettextitems getTextItems�� 0 trim  
�� .aevtoappnull  �   � ****� ������������� 0 	csvtolist 	csvToList�� ����� �  ������ 0 csvtext csvText�� 0 implementation  ��  � ���������������������������������� 0 csvtext csvText�� 0 implementation  �� 0 	separator  �� 0 trimming  �� 0 o  �� 	0 astid  �� 0 	qdticount 	qdtiCount�� "0 quoteinprogress quoteInProgress�� 0 i  �� 0 thisbit thisBit�� 0 a  �� 0 	thisfield 	thisField�� (0 possiblefieldcount possibleFieldCount�� 0 j  �� 0 thesefields theseFields�� 0 k  � �������������������1����q���������������������� 0 	separator  �� 0 trimming  �� �� 0 o  � �����������
�� .ascrinit****      � ****� k     �� ��� ��� ��� �����  ��  ��  � ���������� 0 qdti  �� 0 currentrecord currentRecord��  0 possiblefields possibleFields�� 0 
recordlist 
recordList� ��������������� 0 gettextitems getTextItems�� 0 qdti  �� 0 currentrecord currentRecord
�� 
msng��  0 possiblefields possibleFields�� 0 
recordlist 
recordList�� *b   �l+ �Ojv�O�Ojv�
�� 
ascr
�� 
txdl�� 0 qdti  
�� .corecnte****       ****
�� 
cobj
�� 
bool
�� 
TEXT
�� 
citm�� 0 currentrecord currentRecord�� 0 gettextitems getTextItems��  0 possiblefields possibleFields
�� 
kocl
�� 
cpar�� 0 trim  �� 0 
recordlist 
recordList��=����f�%E[�,E�Z[�,E�ZO��K S�O��,E�O��,j 
E�OfE�O�g��k�lh ��,�/E�O�j 
j
 �� �&�� E���,FO��,[�\[Z�\Z�k2�&E�Oa ��,FO�a -E�Oa ��,FO��&�a ,6FOfE�Y �k ��,�k/�a ,6FY hO*��l+ �a ,FO�a ,j 
E�Ok�kh �a ,�/E�O�a a l 
k ��a -E�O kk�j 
kkh ��/E�O�k
 �k�&
 �k �&
 *�el+ j 
j�& *��l+ �a ,6FY hO�a ,�a ,6FOjv�a ,F[OY��O�6E�O��
 �j 
j�& *��l+ �a ,6FY hY L�k	 ��
 �� �&�&
 �k 	 �k �&�&
 *�el+ j 
j�& *��l+ �a ,6FY h[OY��OPY � hY �k �kE�OeE�Y h[OY�/VO�a ,jv �a ,�a ,6FY hO���,FO�a ,E� ������������ 0 gettextitems getTextItems�� ����� �  ������ 0 txt  �� 	0 delim  ��  � ���������������� 0 txt  �� 	0 delim  �� 	0 astid  �� 0 ticount tiCount�� 0 	textitems 	textItems�� 0 i  �� 0 j  � ������������
�� 
ascr
�� 
txdl
�� 
citm
�� .corecnte****       ****�������� W��,E�O���,FO��-j E�OjvE�O 1k��h ��E�O�� �E�Y hO��[�\[Z�\Z�2%E�[OY��O���,FO�� ��`���������� 0 trim  �� ����� �  ������ 0 txt  �� 0 trimming  ��  � �������� 0 txt  �� 0 trimming  �� 0 i  � ���������
�� .corecnte****       ****
�� 
spac
�� 
ctxt������ s� l +k�j  kkh �� �[�\[Zl\Zi2E�Y [OY��O +k�j  kkh �� �[�\[Zk\Z�2E�Y [OY��O��  �E�Y hY hO�� �����������
�� .aevtoappnull  �   � ****� k    ��   ��  '��  +��  <��  B��  U��  _��  i��  q��  z��  ���  ���  �����  ��  ��  � ���� 0 
therequest 
theRequest� D������~ 8�}�|�{�z�y N�x�w�v�u�t�s�r g�q�p�o�n�m�l�k�j�i�h�g�f�e�d�c�b�a�`�_�^�]��\�[�Z�Y ��X�W�V�U�T�S)+9;GI[giu���
�� .earsffdralis        afdr�� "0 containerfolder containerFolder
� .ascrcmnt****      � ****
�~ 
psxp
�} 
psxf�| 0 	batchpath 	batchPath
�{ .rdwrread****        ****�z 0 csvtext csvText�y 0 	separator  �x 0 trimming  �w 0 	csvtolist 	csvToList�v  0 listofrequests listOfRequests
�u 
cobj
�t .corecnte****       ****�s 0 numberofcols numberOfCols�r 
�q  0 requestheaders requestHeaders�p .0 listofrequestsnoheads listOfRequestsNoHeads�o $0 numberofrequests numberOfRequests
�n 
kocl�m 0 maintraveller mainTraveller�l 0 travelsummary travelSummary�k 0 fromdate fromDate�j �i 0 todate toDate�h �g 0 activity  �f �e 
0 reason  �d 
�c 
bool�b 0 flightbaggage flightBaggage�a �` 0 
flightneed 
flightNeed�_ 	�^ 0 	hotelneed 	hotelNeed�] 0 	trainneed 	trainNeed
�\ 
cwin
�[ 
bTab
�Z 
prdt
�Y 
pURL
�X .corecrel****      � null
�W 
cTab
�V .sysodelanull��� ��� nmbr
�U 
dcnm
�T 
docu
�S .sfridojs****       utxt���)j  E�O�j O��,�%�&E�O�j E�O*���l�elm+ E�O��k/j E` O_ a  )ja Y hO��k/E` O�[�\[Zl\Zi2E` O_ j E` O_ j O_ j O�_ [a �l kh  ��k/E` O��l/E` O��m/E` O��a /E` O��a /E` O��a /E` O��a  /a !&E` "O��a #/a !&E` $O��a %/a !&E` &O��a /a !&E` 'Oa ([*a )k/  *a a *a +a ,a -la  .*a /,FUOmj 0Oa 1a 2*a 3k/l 4Omj 0Oa 5_ %a 6%a 2*a 3k/l 4Oa 7_ %a 8%a 2*a 3k/l 4Oa 9_ %a :%a 2*a 3k/l 4Oa ;_ %a <%a 2*a 3k/l 4Omj 0Oa =a 2*a 3k/l 4Oa >_ %a ?%a 2*a 3k/l 4O_ "e  a @a 2*a 3k/l 4Y hO_ $ a Aa 2*a 3k/l 4Olj 0Y hO_ & a Ba 2*a 3k/l 4Olj 0Y hO_ ' a Ca 2*a 3k/l 4Olj 0Y hUOmj 0[OY�ascr  ��ޭ