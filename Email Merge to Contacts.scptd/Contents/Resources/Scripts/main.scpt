FasdUAS 1.101.10   ��   ��    k             l     ��  ��    W Q --------------------------------------------------------------------------------     � 	 	 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -   
  
 l     ��  ��    : 4 Author:	Craig Buchanan <craig.buchanan@cogniza.com>     �   h   A u t h o r : 	 C r a i g   B u c h a n a n   < c r a i g . b u c h a n a n @ c o g n i z a . c o m >      l     ��  ��      Created:	2013-09-07     �   (   C r e a t e d : 	 2 0 1 3 - 0 9 - 0 7      l     ��  ��    | v Purpose:	Send a personalize email message to each Contact in the selected OS X Address Book group or Outlook Category     �   �   P u r p o s e : 	 S e n d   a   p e r s o n a l i z e   e m a i l   m e s s a g e   t o   e a c h   C o n t a c t   i n   t h e   s e l e c t e d   O S   X   A d d r e s s   B o o k   g r o u p   o r   O u t l o o k   C a t e g o r y      l     ��  ��      Installation:      �      I n s t a l l a t i o n :        l     ��   !��     T N			1) Drag bundle (Email Merge to Contacts.scptd) to ~/Library/Scripts/ folder    ! � " " � 	 	 	 1 )   D r a g   b u n d l e   ( E m a i l   M e r g e   t o   C o n t a c t s . s c p t d )   t o   ~ / L i b r a r y / S c r i p t s /   f o l d e r   # $ # l     �� % &��   %  	 Usage:		    & � ' '    U s a g e : 	 	 $  ( ) ( l     �� * +��   * B <			1) Select AppleScrip menu, then 'Email Merge to Contacts'    + � , , x 	 	 	 1 )   S e l e c t   A p p l e S c r i p   m e n u ,   t h e n   ' E m a i l   M e r g e   t o   C o n t a c t s ' )  - . - l     �� / 0��   /   Todos:    0 � 1 1    T o d o s : .  2 3 2 l     �� 4 5��   4 % 			1) Handle (Mail) attachments    5 � 6 6 > 	 	 	 1 )   H a n d l e   ( M a i l )   a t t a c h m e n t s 3  7 8 7 l     �� 9 :��   9 4 .			2) Save 'theResult' variables to a CSV file    : � ; ; \ 	 	 	 2 )   S a v e   ' t h e R e s u l t '   v a r i a b l e s   t o   a   C S V   f i l e 8  < = < l     �� > ?��   > l f			3) Convert theVariables to a Record and adjust dependent script accordingly; move to helper script?    ? � @ @ � 	 	 	 3 )   C o n v e r t   t h e V a r i a b l e s   t o   a   R e c o r d   a n d   a d j u s t   d e p e n d e n t   s c r i p t   a c c o r d i n g l y ;   m o v e   t o   h e l p e r   s c r i p t ? =  A B A l     �� C D��   C W Q --------------------------------------------------------------------------------    D � E E �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - B  F G F l     ��������  ��  ��   G  H I H l     �� J K��   J W Q --------------------------------------------------------------------------------    K � L L �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - I  M N M l     �� O P��   O #  default string localizations    P � Q Q :   d e f a u l t   s t r i n g   l o c a l i z a t i o n s N  R S R l     �� T U��   T W Q --------------------------------------------------------------------------------    U � V V �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - S  W X W l     Y���� Y r      Z [ Z m      \ \ � ] ]  E m a i l   M e r g e [ o      ���� 0 dialog_title DIALOG_TITLE��  ��   X  ^ _ ^ l     ��������  ��  ��   _  ` a ` l    b���� b r     c d c m     e e � f f  O K d o      ���� 0 	button_ok 	BUTTON_OK��  ��   a  g h g l    i���� i r     j k j m    	 l l � m m  C a n c e l k o      ���� 0 button_cancel BUTTON_CANCEL��  ��   h  n o n l    p���� p r     q r q m     s s � t t  Y e s r o      ���� 0 
button_yes 
BUTTON_YES��  ��   o  u v u l    w���� w r     x y x m     z z � { {  N o y o      ���� 0 	button_no 	BUTTON_NO��  ��   v  | } | l     ��������  ��  ��   }  ~  ~ l    ����� � r     � � � m     � � � � �  M a i l � o      ���� 0 button_mail BUTTON_MAIL��  ��     � � � l    ����� � r     � � � m     � � � � �  O u t l o o k � o      ����  0 button_outlook BUTTON_OUTLOOK��  ��   �  � � � l     ��������  ��  ��   �  � � � l    ����� � r     � � � m     � � � � � F D o   y o u   w a n t   t o   r u n   i n   t e s t i n g   m o d e ? � o      ���� *0 prompt_testing_mode PROMPT_TESTING_MODE��  ��   �  � � � l    ' ����� � r     ' � � � m     # � � � � � * S e n d   a l l   m e s s a g e s   t o : � o      ���� *0 prompt_what_address PROMPT_WHAT_ADDRESS��  ��   �  � � � l  ( / ����� � r   ( / � � � m   ( + � � � � � 4 e m a i l . m e r g e @ m a i l i n a t o r . c o m � o      ���� 0 answer_email ANSWER_EMAIL��  ��   �  � � � l     ��������  ��  ��   �  � � � l  0 7 ����� � r   0 7 � � � m   0 3 � � � � �  W h i c h   p r o v i d e r ? � o      ���� .0 prompt_which_provider PROMPT_WHICH_PROVIDER��  ��   �  � � � l  8 ? ����� � r   8 ? � � � m   8 ; � � � � � 4 W h i c h   g r o u p   s h a l l   b e   u s e d ? � o      ���� (0 prompt_which_group PROMPT_WHICH_GROUP��  ��   �  � � � l  @ G ����� � r   @ G � � � m   @ C � � � � � r T h e   g r o u p   o f   c o n t a c t s   c o n t a i n e d   n o   v a l i d   e m a i l   a d d r e s s e s . � o      ���� 00 prompt_no_valid_emails PROMPT_NO_VALID_EMAILS��  ��   �  � � � l  H O ����� � r   H O � � � m   H K � � � � � F W h i c h   e m a i l - a d d r e s s   t y p e ( s )   t o   u s e ? � o      ���� 60 prompt_which_address_type PROMPT_WHICH_ADDRESS_TYPE��  ��   �  � � � l  P W ����� � r   P W � � � m   P S � � � � � , W h i c h   t e m p l a t e   t o   u s e ? � o      ���� .0 prompt_which_template PROMPT_WHICH_TEMPLATE��  ��   �  � � � l  X _ ����� � r   X _ � � � m   X [ � � � � � � P a u s i n g   b e t w e e n   b a t c h e s .     P r o c e s s i n g   w i l l   r e s u m e   a t   % s .   ( t h i s   a l e r t   w i l l   b e   a u t o m a t i c a l l y   d i s m i s s e d ) � o      ���� 0 promt_pausing PROMT_PAUSING��  ��   �  � � � l  ` g ����� � r   ` g � � � m   ` c � � � � � F % d   e m a i l s   w e r e   s e n t   t o   % d   c o n t a c t s . � o      ���� 0 prompt_done PROMPT_DONE��  ��   �  � � � l     ��������  ��  ��   �  � � � l  h o ����� � r   h o � � � m   h k � � � � � $ N o   e m a i l   a d d r e s s e s � o      ���� &0 reason_no_address REASON_NO_ADDRESS��  ��   �  � � � l  p w ����� � r   p w � � � m   p s � � � � � " D u p l i c a t e   a d d r e s s � o      ���� 40 reason_duplicate_address REASON_DUPLICATE_ADDRESS��  ��   �  � � � l  x  ����� � r   x  � � � m   x { � � � � � $ W r o n g   a d d r e s s   t y p e � o      ���� 60 reason_wrong_address_type REASON_WRONG_ADDRESS_TYPE��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � W Q --------------------------------------------------------------------------------    � � � � �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     �� � ���   �  	 settings    � � � �    s e t t i n g s �  � � � l     �� � ���   � W Q --------------------------------------------------------------------------------    � � � � �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �    l  � ����� r   � � J   � �  m   � � �		  % % F i r s t N a m e % % 

 m   � � �  % % L a s t N a m e % %  m   � � �  % % D i s p l a y N a m e % %  m   � � �  % % C o m p a n y % %  m   � � �   % % E m a i l A d d r e s s % %  m   � � �  % % D a t e % % �� m   � � �    % % T i m e % %��   o      ���� 0 thevariables theVariables��  ��   !"! l     ��#$��  # 6 0 prevent messages from being perceived as 'spam'   $ �%% `   p r e v e n t   m e s s a g e s   f r o m   b e i n g   p e r c e i v e d   a s   ' s p a m '" &'& l  � �()*( r   � �+,+ m   � ����� 2, o      ���� 0 	batchsize 	batchSize) = 7 # of messages to deliver sequentially; recommended: 50   * �-- n   #   o f   m e s s a g e s   t o   d e l i v e r   s e q u e n t i a l l y ;   r e c o m m e n d e d :   5 0' ./. l  � �0120 r   � �343 m   � ����� 4 o      ���� 0 thepause thePause1 3 - # of minutes between batches; recommended: 5   2 �55 Z   #   o f   m i n u t e s   b e t w e e n   b a t c h e s ;   r e c o m m e n d e d :   5/ 676 l     ��������  ��  ��  7 898 l     ��:;��  : W Q --------------------------------------------------------------------------------   ; �<< �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -9 =>= l     ��?@��  ?   external libraries   @ �AA &   e x t e r n a l   l i b r a r i e s> BCB l     ��DE��  D W Q --------------------------------------------------------------------------------   E �FF �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -C GHG l     ��������  ��  ��  H IJI l     ��KL��  K K Eset userHome to (POSIX path of (path to home folder as text) as text)   L �MM � s e t   u s e r H o m e   t o   ( P O S I X   p a t h   o f   ( p a t h   t o   h o m e   f o l d e r   a s   t e x t )   a s   t e x t )J NON l     ��PQ��  P 9 3set userScriptHome to userHome & "Library/Scripts/"   Q �RR f s e t   u s e r S c r i p t H o m e   t o   u s e r H o m e   &   " L i b r a r y / S c r i p t s / "O STS l     ��UV��  U e _set outlookHome to userHome & "Application Support/Microsoft/Office/Outlook Script Menu Items/"   V �WW � s e t   o u t l o o k H o m e   t o   u s e r H o m e   &   " A p p l i c a t i o n   S u p p o r t / M i c r o s o f t / O f f i c e / O u t l o o k   S c r i p t   M e n u   I t e m s / "T XYX l     ��������  ��  ��  Y Z[Z l     �\]�  \ a [property Helper : load script (path to resource "HelperLibrary.scpt" in directory "Scripts"   ] �^^ � p r o p e r t y   H e l p e r   :   l o a d   s c r i p t   ( p a t h   t o   r e s o u r c e   " H e l p e r L i b r a r y . s c p t "   i n   d i r e c t o r y   " S c r i p t s "[ _`_ l  � �a�~�}a r   � �bcb I  � ��|d�{
�| .sysoloadscpt        filed l  � �e�z�ye I  � ��xfg
�x .sysorpthalis        TEXTf m   � �hh �ii $ H e l p e r L i b r a r y . s c p tg �wj�v
�w 
in Dj m   � �kk �ll  S c r i p t s�v  �z  �y  �{  c o      �u�u 0 helper Helper�~  �}  ` mnm l     �t�s�r�t  �s  �r  n opo l     �qqr�q  q W Q --------------------------------------------------------------------------------   r �ss �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -p tut l     �pvw�p  v   choose plugin   w �xx    c h o o s e   p l u g i nu yzy l     �o{|�o  { W Q --------------------------------------------------------------------------------   | �}} �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -z ~~ l  � ���n�m� I  � ��l��
�l .sysodlogaskr        TEXT� o   � ��k�k .0 prompt_which_provider PROMPT_WHICH_PROVIDER� �j��
�j 
appr� o   � ��i�i 0 dialog_title DIALOG_TITLE� �h��
�h 
btns� J   � ��� ��� o   � ��g�g 0 button_mail BUTTON_MAIL� ��f� o   � ��e�e  0 button_outlook BUTTON_OUTLOOK�f  � �d��c
�d 
dflt� o   � ��b�b 0 button_mail BUTTON_MAIL�c  �n  �m   ��� l     �a�`�_�a  �`  �_  � ��� l  �*��^�]� Z   �*����\� =   � ���� 1   � ��[
�[ 
rslt� K   � ��� �Z��Y
�Z 
bhit� o   � ��X�X 0 button_mail BUTTON_MAIL�Y  � k   � ��� ��� r   � ���� I  � ��W��V
�W .sysoloadscpt        file� l  � ���U�T� I  � ��S��
�S .sysorpthalis        TEXT� m   � ��� ���  M a i l P l u g i n . s c p t� �R��Q
�R 
in D� m   � ��� ���  S c r i p t s�Q  �U  �T  �V  � o      �P�P 0 	theplugin 	thePlugin� ��O� l  � ��N�M�L�N  �M  �L  �O  � ��� =  ��� 1  �K
�K 
rslt� K  �� �J��I
�J 
bhit� o  	�H�H  0 button_outlook BUTTON_OUTLOOK�I  � ��G� k  &�� ��� r  $��� I  �F��E
�F .sysoloadscpt        file� l ��D�C� I �B��
�B .sysorpthalis        TEXT� m  �� ��� $ O u t l o o k P l u g i n . s c p t� �A��@
�A 
in D� m  �� ���  S c r i p t s�@  �D  �C  �E  � o      �?�? 0 	theplugin 	thePlugin� ��>� l %%�=�<�;�=  �<  �;  �>  �G  �\  �^  �]  � ��� l     �:�9�8�:  �9  �8  � ��� l     �7���7  � W Q --------------------------------------------------------------------------------   � ��� �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �6���6  �   testing mode   � ���    t e s t i n g   m o d e� ��� l     �5���5  � W Q --------------------------------------------------------------------------------   � ��� �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �4�3�2�4  �3  �2  � ��� l +E��1�0� r  +E��� l +A��/�.� I +A�-��
�- .sysodlogaskr        TEXT� o  +,�,�, *0 prompt_testing_mode PROMPT_TESTING_MODE� �+��
�+ 
appr� o  /0�*�* 0 dialog_title DIALOG_TITLE� �)��
�) 
btns� J  37�� ��� o  34�(�( 0 
button_yes 
BUTTON_YES� ��'� o  45�&�& 0 	button_no 	BUTTON_NO�'  � �%��$
�% 
dflt� o  :;�#�# 0 
button_yes 
BUTTON_YES�$  �/  �.  � o      �"�" 0 testingmode testingMode�1  �0  � ��� l F���!� � Z  F������ =  FP��� 1  FI�
� 
rslt� K  IO�� ���
� 
bhit� o  LM�� 0 	button_no 	BUTTON_NO�  � k  SZ�� ��� r  SX��� m  ST�
� boovfals� o      �� 0 testingmode testingMode� ��� l YY����  �  �  �  � ��� =  ]g��� 1  ]`�
� 
rslt� K  `f�� ���
� 
bhit� o  cd�� 0 
button_yes 
BUTTON_YES�  � ��� k  j��� ��� r  jo��� m  jk�
� boovtrue� o      �� 0 testingmode testingMode� ��� l pp����  �  �  � � � r  p� n  p� 1  ���

�
 
ttxt l p��	� I p��
� .sysodlogaskr        TEXT o  ps�� *0 prompt_what_address PROMPT_WHAT_ADDRESS �	
� 
dtxt o  vy�� 0 answer_email ANSWER_EMAIL	 �

� 
appr
 o  |}�� 0 dialog_title DIALOG_TITLE �
� 
btns J  �� �  o  ������ 0 	button_ok 	BUTTON_OK�    ����
�� 
dflt o  ������ 0 	button_ok 	BUTTON_OK��  �	  �   o      ����  0 testingaddress testingAddress   Z ������ = �� 1  ����
�� 
rslt m  ����
�� boovfals L  ������  ��  ��   �� l ����������  ��  ��  ��  �  �  �!  �   �  l     ��������  ��  ��    l     ����   W Q --------------------------------------------------------------------------------    � �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -   l     ��!"��  !   select group/contacts   " �## ,   s e l e c t   g r o u p / c o n t a c t s  $%$ l     ��&'��  & W Q --------------------------------------------------------------------------------   ' �(( �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -% )*) l     ��������  ��  ��  * +,+ l     ��-.��  - $ generate a List of group names   . �// < g e n e r a t e   a   L i s t   o f   g r o u p   n a m e s, 010 l ��2����2 r  ��343 n ��565 I  ���������� 0 	getgroups 	getGroups��  ��  6 o  ������ 0 	theplugin 	thePlugin4 o      ���� 0 
groupnames 
groupNames��  ��  1 787 l     ��������  ��  ��  8 9:9 l     ��;<��  ; #  sort the list alphabetically   < �== :   s o r t   t h e   l i s t   a l p h a b e t i c a l l y: >?> l ��@����@ r  ��ABA n ��CDC I  ����E���� 0 simple_sort  E F��F l ��G����G o  ������ 0 
groupnames 
groupNames��  ��  ��  ��  D o  ������ 0 helper HelperB o      ���� 0 
groupnames 
groupNames��  ��  ? HIH l     ��������  ��  ��  I JKJ l     ��LM��  L 2 , prompt to select a group name from the List   M �NN X   p r o m p t   t o   s e l e c t   a   g r o u p   n a m e   f r o m   t h e   L i s tK OPO l ��Q����Q r  ��RSR l ��T����T I ����UV
�� .gtqpchltns    @   @ ns  U o  ������ 0 
groupnames 
groupNamesV ��WX
�� 
apprW o  ������ 0 dialog_title DIALOG_TITLEX ��Y��
�� 
prmpY o  ������ (0 prompt_which_group PROMPT_WHICH_GROUP��  ��  ��  S o      ���� 0 	groupname 	groupName��  ��  P Z[Z l ��\����\ Z ��]^����] = ��_`_ 1  ����
�� 
rslt` m  ����
�� boovfals^ L  ������  ��  ��  ��  ��  [ aba l     ��������  ��  ��  b cdc l ��e����e r  ��fgf n ��hih I  ����j���� 0 getcontacts getContactsj k��k o  ������ 0 	groupname 	groupName��  ��  i o  ������ 0 	theplugin 	thePluging o      ���� 0 thecontacts theContacts��  ��  d lml l �
n����n Z �
op����o = �qrq 1  ����
�� 
rsltr m  � ��
�� boovfalsp L  ����  ��  ��  ��  ��  m sts l     ��������  ��  ��  t uvu l     ��wx��  w E ? promtp to select which email-address type (e.g. 'work') to use   x �yy ~   p r o m t p   t o   s e l e c t   w h i c h   e m a i l - a d d r e s s   t y p e   ( e . g .   ' w o r k ' )   t o   u s ev z{z l 9|����| Z  9}~����} =  � l ������ I �����
�� .corecnte****       ****� n  ��� o  ���� 0 thetypes theTypes� o  ���� 0 thecontacts theContacts��  ��  ��  � m  ����  ~ k  5�� ��� I 2����
�� .sysodlogaskr        TEXT� o  ���� 00 prompt_no_valid_emails PROMPT_NO_VALID_EMAILS� ����
�� 
appr� o  !"���� 0 dialog_title DIALOG_TITLE� ����
�� 
btns� J  %(�� ���� o  %&���� 0 	button_ok 	BUTTON_OK��  � �����
�� 
dflt� o  +,���� 0 	button_ok 	BUTTON_OK��  � ���� L  35����  ��  ��  ��  ��  ��  { ��� l     ��������  ��  ��  � ��� l :Y������ r  :Y��� l :U������ I :U����
�� .gtqpchltns    @   @ ns  � l :A������ n  :A��� o  =A���� 0 thetypes theTypes� o  :=���� 0 thecontacts theContacts��  ��  � ����
�� 
appr� o  DE���� 0 dialog_title DIALOG_TITLE� ����
�� 
prmp� o  HK�� 60 prompt_which_address_type PROMPT_WHICH_ADDRESS_TYPE� �~��}
�~ 
mlsl� m  NO�|
�| boovtrue�}  ��  ��  � o      �{�{ 0 thetypes theTypes��  ��  � ��� l Zh��z�y� Z Zh���x�w� = Z_��� 1  Z]�v
�v 
rslt� m  ]^�u
�u boovfals� L  bd�t�t  �x  �w  �z  �y  � ��� l     �s�r�q�s  �r  �q  � ��� l     �p���p  � W Q --------------------------------------------------------------------------------   � ��� �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �o���o  �   choose template to send   � ��� 0   c h o o s e   t e m p l a t e   t o   s e n d� ��� l     �n���n  � W Q --------------------------------------------------------------------------------   � ��� �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �m�l�k�m  �l  �k  � ��� l     �j���j  � 7 1 generate a List of drafts in the 'Drafts' folder   � ��� b   g e n e r a t e   a   L i s t   o f   d r a f t s   i n   t h e   ' D r a f t s '   f o l d e r� ��� l it��i�h� r  it��� n ip��� I  lp�g�f�e�g 0 gettemplates getTemplates�f  �e  � o  il�d�d 0 	theplugin 	thePlugin� o      �c�c 0 thesubjectids theSubjectIds�i  �h  � ��� l     �b�a�`�b  �a  �`  � ��� l     �_���_  � #  sort the list alphabetically   � ��� :   s o r t   t h e   l i s t   a l p h a b e t i c a l l y� ��� l u���^�]� r  u���� n u��� I  x�\��[�\ 0 simple_sort  � ��Z� l x{��Y�X� o  x{�W�W 0 thesubjectids theSubjectIds�Y  �X  �Z  �[  � o  ux�V�V 0 helper Helper� o      �U�U 0 thesubjectids theSubjectIds�^  �]  � ��� l     �T�S�R�T  �S  �R  � ��� l     �Q���Q  � 2 , prompt to select a group name from the List   � ��� X   p r o m p t   t o   s e l e c t   a   g r o u p   n a m e   f r o m   t h e   L i s t� ��� l ����P�O� r  ����� l ����N�M� I ���L��
�L .gtqpchltns    @   @ ns  � o  ���K�K 0 thesubjectids theSubjectIds� �J��
�J 
appr� o  ���I�I 0 dialog_title DIALOG_TITLE� �H��G
�H 
prmp� o  ���F�F .0 prompt_which_template PROMPT_WHICH_TEMPLATE�G  �N  �M  � o      �E�E 0 	thechoice 	theChoice�P  �O  � ��� l ����D�C� Z �����B�A� = ����� 1  ���@
�@ 
rslt� m  ���?
�? boovfals� L  ���>�>  �B  �A  �D  �C  � ��� l     �=�<�;�=  �<  �;  � ��� l     �:���:  �   extract ID from string   � ��� .   e x t r a c t   I D   f r o m   s t r i n g� ��� l ����9�8� r  ����� n ��   I  ���7�6�7 0 find    m  �� �  [ �5 o  ���4�4 0 	thechoice 	theChoice�5  �6   o  ���3�3 0 helper Helper� o      �2�2 0 n  �9  �8  � 	 l ��
�1�0
 r  �� n �� I  ���/�.�/ 0 find    m  �� �  ] �- o  ���,�, 0 	thechoice 	theChoice�-  �.   o  ���+�+ 0 helper Helper o      �*�* 0 m  �1  �0  	  l ���)�( r  �� c  �� l ���'�& n  �� 7���% 
�% 
ctxt l ��!�$�#! [  ��"#" o  ���"�" 0 n  # m  ���!�! �$  �#    l ��$� �$ \  ��%&% o  ���� 0 m  & m  ���� �   �   l ��'��' c  ��()( o  ���� 0 	thechoice 	theChoice) m  ���
� 
ctxt�  �  �'  �&   m  ���
� 
ctxt o      �� 0 theid theID�)  �(   *+* l     ����  �  �  + ,-, l     �./�  . W Q --------------------------------------------------------------------------------   / �00 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- 121 l     �34�  3   iterate and send   4 �55 "   i t e r a t e   a n d   s e n d2 676 l     �89�  8 W Q --------------------------------------------------------------------------------   9 �:: �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -7 ;<; l     ����  �  �  < =>= l ��?��? r  ��@A@ l ��B��
B I ���	��
�	 .misccurdldt    ��� null�  �  �  �
  A o      �� 0 	starttime 	startTime�  �  > CDC l �EFGE r  �HIH m  ����  I o      �� 0 iterator  F   tracks batch size   G �JJ $   t r a c k s   b a t c h   s i z eD KLK l     ����  �  �  L MNM l OPQO r  RSR m  � �   S o      ���� $0 recordsprocessed recordsProcessedP   the number of contacts   Q �TT .   t h e   n u m b e r   o f   c o n t a c t sN UVU l 	WXYW r  	Z[Z m  	
����  [ o      ���� 0 messagessent messagesSentX + % the number of messages actually sent   Y �\\ J   t h e   n u m b e r   o f   m e s s a g e s   a c t u a l l y   s e n tV ]^] l     ��������  ��  ��  ^ _`_ l abca r  ded J  ����  e o      ���� .0 theaddressesprocessed theAddressesProcessedb   track email address   c �ff (   t r a c k   e m a i l   a d d r e s s` ghg l ijki r  lml J  ����  m o      ���� 0 
theresults 
theResultsj   track results   k �nn    t r a c k   r e s u l t sh opo l     ��������  ��  ��  p qrq l �stus X  �v��wv l 7�xyzx U  7�{|{ l >�}~} k  >��� ��� l >>��������  ��  ��  � ��� r  >G��� [  >C��� o  >A���� $0 recordsprocessed recordsProcessed� m  AB���� � o      ���� $0 recordsprocessed recordsProcessed� ��� l HH��������  ��  ��  � ��� l HH������  � < 6 if the contact doesn't have any email addresses, skip   � ��� l   i f   t h e   c o n t a c t   d o e s n ' t   h a v e   a n y   e m a i l   a d d r e s s e s ,   s k i p� ��� Z  H�������� =  HS��� n  HQ��� 1  MQ��
�� 
leng� l HM������ n  HM��� o  IM����  0 emailaddresses emailAddresses� o  HI���� 0 
thecontact 
theContact��  ��  � m  QR����  � k  V��� ��� l VV��������  ��  ��  � ��� r  V���� K  V��� ������ 0 displayname displayName� l Y^������ n  Y^��� o  Z^���� 0 displayname displayName� o  YZ���� 0 
thecontact 
theContact��  ��  � ������ 0 companyname companyName� l af������ n  af��� o  bf���� 0 companyname companyName� o  ab���� 0 
thecontact 
theContact��  ��  � ������ 0 emailaddress emailAddress� m  il��
�� 
null� ������ 0 type  � m  or��
�� 
null� ������ 0 sent  � m  uv��
�� boovfals� ������� 
0 reason  � o  y|���� &0 reason_no_address REASON_NO_ADDRESS��  � o      ���� 0 	theresult 	theResult� ��� l ����������  ��  ��  � ��� l ��������  � x rdisplay dialog Helper's format_string("%s - %s: %s", {(displayName of theContact), "NOT SENT", REASON_NO_ADDRESS})   � ��� � d i s p l a y   d i a l o g   H e l p e r ' s   f o r m a t _ s t r i n g ( " % s   -   % s :   % s " ,   { ( d i s p l a y N a m e   o f   t h e C o n t a c t ) ,   " N O T   S E N T " ,   R E A S O N _ N O _ A D D R E S S } )� ��� l ����������  ��  ��  � ��� l ������  S  ���  loop 0   � ���  l o o p   0� ���� l ����������  ��  ��  ��  ��  ��  � ��� l ����������  ��  ��  � ��� l ��������  � + % otherwise process each email address   � ��� J   o t h e r w i s e   p r o c e s s   e a c h   e m a i l   a d d r e s s� ��� l ������ X  ������� l ������ U  ����� l ������ k  ���� ��� l ����������  ��  ��  � ��� l ��������  � M G skip if the address type (e.g. 'work') is not one of the chosen values   � ��� �   s k i p   i f   t h e   a d d r e s s   t y p e   ( e . g .   ' w o r k ' )   i s   n o t   o n e   o f   t h e   c h o s e n   v a l u e s� ��� Z  �8������ H  ���� E  ����� o  ������ 0 thetypes theTypes� l �������� n  ����� o  ������ 0 thetype theType� o  ������ "0 theemailaddress theEmailAddress��  ��  � k  ���� ��� l ����������  ��  ��  � ��� r  ����� K  ���� ������ 0 displayname displayName� l �������� n  ����� o  ������ 0 displayname displayName� o  ������ 0 
thecontact 
theContact��  ��  � ������ 0 companyname companyName� l �������� n  ����� o  ������ 0 companyname companyName� o  ������ 0 
thecontact 
theContact��  ��  � ��� �� 0 emailaddress emailAddress� l ������ n  �� o  ������ 0 
theaddress 
theAddress o  ������ "0 theemailaddress theEmailAddress��  ��    ���� 0 type   l ������ n  �� o  ������ 0 thetype theType o  ������ "0 theemailaddress theEmailAddress��  ��   ��	
�� 0 sent  	 m  ����
�� boovfals
 ������ 
0 reason   o  ������ 60 reason_wrong_address_type REASON_WRONG_ADDRESS_TYPE��  � o      ���� 0 	theresult 	theResult�  l ����������  ��  ��    l ������   � �display dialog Helper's format_string("%s (%s:%s) - %s: %s", {(displayName of theContact), (theType of theEmailAddress), (theAddress of theEmailAddress), "NOT SENT", REASON_WRONG_ADDRESS_TYPE})    �� d i s p l a y   d i a l o g   H e l p e r ' s   f o r m a t _ s t r i n g ( " % s   ( % s : % s )   -   % s :   % s " ,   { ( d i s p l a y N a m e   o f   t h e C o n t a c t ) ,   ( t h e T y p e   o f   t h e E m a i l A d d r e s s ) ,   ( t h e A d d r e s s   o f   t h e E m a i l A d d r e s s ) ,   " N O T   S E N T " ,   R E A S O N _ W R O N G _ A D D R E S S _ T Y P E } )  l ����������  ��  ��    l ��  S  ��  loop 1    �  l o o p   1  l ����������  ��  ��   �� l ������   ) # skip if the address is a duplicate    �   F   s k i p   i f   t h e   a d d r e s s   i s   a   d u p l i c a t e��  � !"! E  ��#$# o  ������ .0 theaddressesprocessed theAddressesProcessed$ l ��%����% n  ��&'& o  ������ 0 
theaddress 
theAddress' o  ������ "0 theemailaddress theEmailAddress��  ��  " (��( k  �4)) *+* l ����~�}�  �~  �}  + ,-, r  �0./. K  �,00 �|12�| 0 displayname displayName1 l 3�{�z3 n  454 o  �y�y 0 displayname displayName5 o  �x�x 0 
thecontact 
theContact�{  �z  2 �w67�w 0 companyname companyName6 l 	8�v�u8 n  	9:9 o  
�t�t 0 companyname companyName: o  	
�s�s 0 
thecontact 
theContact�v  �u  7 �r;<�r 0 emailaddress emailAddress; l =�q�p= n  >?> o  �o�o 0 
theaddress 
theAddress? o  �n�n "0 theemailaddress theEmailAddress�q  �p  < �m@A�m 0 type  @ l B�l�kB n  CDC o  �j�j 0 thetype theTypeD o  �i�i "0 theemailaddress theEmailAddress�l  �k  A �hEF�h 0 sent  E m  !"�g
�g boovfalsF �fG�e�f 
0 reason  G o  %(�d�d 40 reason_duplicate_address REASON_DUPLICATE_ADDRESS�e  / o      �c�c 0 	theresult 	theResult- HIH l 11�b�a�`�b  �a  �`  I JKJ l 11�_LM�_  L � �display dialog Helper's format_string("%s (%s:%s) - %s: %s", {(displayName of theContact), (theType of theEmailAddress), (theAddress of theEmailAddress), "NOT SENT", REASON_DUPLICATE_ADDRESS})   M �NN� d i s p l a y   d i a l o g   H e l p e r ' s   f o r m a t _ s t r i n g ( " % s   ( % s : % s )   -   % s :   % s " ,   { ( d i s p l a y N a m e   o f   t h e C o n t a c t ) ,   ( t h e T y p e   o f   t h e E m a i l A d d r e s s ) ,   ( t h e A d d r e s s   o f   t h e E m a i l A d d r e s s ) ,   " N O T   S E N T " ,   R E A S O N _ D U P L I C A T E _ A D D R E S S } )K OPO l 11�^�]�\�^  �]  �\  P QRQ l 12STUS  S  12T  loop 1   U �VV  l o o p   1R W�[W l 33�Z�Y�X�Z  �Y  �X  �[  ��  ��  � XYX l 99�W�V�U�W  �V  �U  Y Z[Z l 99�T\]�T  \ ) # copy the message template's fields   ] �^^ F   c o p y   t h e   m e s s a g e   t e m p l a t e ' s   f i e l d s[ _`_ r  9Gaba n 9Ccdc I  <C�Se�R�S 0 
getcontent 
getContente f�Qf o  <?�P�P 0 theid theID�Q  �R  d o  9<�O�O 0 	theplugin 	thePluginb o      �N�N 0 
thecontent 
theContent` ghg l HH�M�L�K�M  �L  �K  h iji l HH�Jkl�J  k   personalize the fields   l �mm .   p e r s o n a l i z e   t h e   f i e l d sj non r  Hapqp n H]rsr I  K]�It�H�I 0 personalize  t uvu o  KN�G�G 0 
thecontent 
theContentv wxw o  NQ�F�F 0 thevariables theVariablesx yzy o  QR�E�E 0 
thecontact 
theContactz {�D{ l RW|�C�B| n  RW}~} o  SW�A�A 0 
theaddress 
theAddress~ o  RS�@�@ "0 theemailaddress theEmailAddress�C  �B  �D  �H  s o  HK�?�? 0 helper Helperq o      �>�> 0 
thecontent 
theContento � l bb�=�<�;�=  �<  �;  � ��� l bb�:���:  � D > create the recipient (use testing email address in test mode)   � ��� |   c r e a t e   t h e   r e c i p i e n t   ( u s e   t e s t i n g   e m a i l   a d d r e s s   i n   t e s t   m o d e )� ��� r  bv��� K  br�� �9���9 0 thename  � m  eh�8
�8 
null� �7��6�7 0 
theaddress 
theAddress� m  kn�5
�5 
null�6  � o      �4�4 0 therecipient theRecipient� ��� Z  w����3�� o  wz�2�2 0 testingmode testingMode� r  }���� K  }��� �1���1 0 thename  � l ����0�/� n  ����� o  ���.�. 0 displayname displayName� o  ���-�- 0 
thecontact 
theContact�0  �/  � �,��+�, 0 
theaddress 
theAddress� o  ���*�*  0 testingaddress testingAddress�+  � o      �)�) 0 therecipient theRecipient�3  � r  ����� K  ���� �(���( 0 thename  � l ����'�&� n  ����� o  ���%�% 0 displayname displayName� o  ���$�$ 0 
thecontact 
theContact�'  �&  � �#��"�# 0 
theaddress 
theAddress� l ����!� � n  ����� o  ���� 0 
theaddress 
theAddress� o  ���� "0 theemailaddress theEmailAddress�!  �   �"  � o      �� 0 therecipient theRecipient� ��� l ������  �  �  � ��� l ������  �   send the message   � ��� "   s e n d   t h e   m e s s a g e� ��� n ����� I  ������ (0 copyandsendmessage copyAndSendMessage� ��� o  ���� 0 theid theID� ��� o  ���� 0 
thecontent 
theContent� ��� o  ���� 0 therecipient theRecipient�  �  � o  ���� 0 	theplugin 	thePlugin� ��� l ������  �  �  � ��� r  ����� K  ���� ���� 0 displayname displayName� l ������ n  ����� o  ���� 0 displayname displayName� o  ���
�
 0 
thecontact 
theContact�  �  � �	���	 0 companyname companyName� l ������ n  ����� o  ���� 0 companyname companyName� o  ���� 0 
thecontact 
theContact�  �  � ���� 0 emailaddress emailAddress� l ������ n  ����� o  ���� 0 
theaddress 
theAddress� o  ��� �  "0 theemailaddress theEmailAddress�  �  � ������ 0 type  � l �������� n  ����� o  ������ 0 thetype theType� o  ������ "0 theemailaddress theEmailAddress��  ��  � ������ 0 sent  � m  ����
�� boovtrue� ������� 
0 reason  � m  ����
�� 
null��  � o      ���� 0 	theresult 	theResult� ��� l ����������  ��  ��  � ��� l ��������  � � �display dialog Helper's format_string("%s (%s:%s) - %s", {(displayName of theContact), (theType of theEmailAddress), (theAddress of theEmailAddress), "SENT"})   � ���< d i s p l a y   d i a l o g   H e l p e r ' s   f o r m a t _ s t r i n g ( " % s   ( % s : % s )   -   % s " ,   { ( d i s p l a y N a m e   o f   t h e C o n t a c t ) ,   ( t h e T y p e   o f   t h e E m a i l A d d r e s s ) ,   ( t h e A d d r e s s   o f   t h e E m a i l A d d r e s s ) ,   " S E N T " } )� ��� l ����������  ��  ��  � ��� l ��������  � !  add address to unique list   � ��� 6   a d d   a d d r e s s   t o   u n i q u e   l i s t� ��� Z  �������� H  ���� E  ����� o  ������ .0 theaddressesprocessed theAddressesProcessed� l �������� n  ����� o  ������ 0 
theaddress 
theAddress� o  ������ "0 theemailaddress theEmailAddress��  ��  � s   
��� l  ������ n   ��� o  ���� 0 
theaddress 
theAddress� o   ���� "0 theemailaddress theEmailAddress��  ��  � l     ������ n      ���  ;  	� o  ���� .0 theaddressesprocessed theAddressesProcessed��  ��  ��  ��  � ��� l ��������  ��  ��  � ��� l ������  �  
 increment   � ���    i n c r e m e n t� ��� r     [   o  ���� 0 iterator   m  ����  o      ���� 0 iterator  �  l ����     update statistics    � $   u p d a t e   s t a t i s t i c s 	
	 r  " [   o  ���� 0 messagessent messagesSent m  ����  o      ���� 0 messagessent messagesSent
  l ##��������  ��  ��    l ##����     manage batches    �    m a n a g e   b a t c h e s  Z  #����� F  #B ?  #4 l #2���� \  #2  l #.!����! I #.��"��
�� .corecnte****       ****" n  #*#$# o  &*���� 0 
therecords 
theRecords$ o  #&���� 0 thecontacts theContacts��  ��  ��    o  .1���� $0 recordsprocessed recordsProcessed��  ��   m  23����   l 7>%����% = 7>&'& o  7:���� 0 iterator  ' o  :=���� 0 	batchsize 	batchSize��  ��   k  E�(( )*) r  EJ+,+ m  EF����  , o      ���� 0 iterator  * -.- l KK��������  ��  ��  . /0/ I K~��12
�� .sysodlogaskr        TEXT1 n Kh343 I  Nh��5���� 0 format_string  5 676 o  NQ���� 0 promt_pausing PROMT_PAUSING7 8��8 J  Qd99 :��: l Qb;����; c  Qb<=< [  Q^>?> l QV@����@ I QV������
�� .misccurdldt    ��� null��  ��  ��  ��  ? l V]A����A ]  V]BCB m  VY���� <C o  Y\���� 0 thepause thePause��  ��  = m  ^a��
�� 
ctxt��  ��  ��  ��  ��  4 o  KN���� 0 helper Helper2 ��DE
�� 
apprD o  kl���� 0 dialog_title DIALOG_TITLEE ��FG
�� 
btnsF J  orHH I��I o  op���� 0 	button_ok 	BUTTON_OK��  G ��J��
�� 
givuJ m  ux���� ��  0 K��K I ���L��
�� .sysodelanull��� ��� nmbrL l �M����M ]  �NON m  ����� <O o  ������ 0 thepause thePause��  ��  ��  ��  ��  ��   P��P l ����������  ��  ��  ��  �   fake loop 1   � �QQ    f a k e   l o o p   1� m  ������ �  /fake 1   � �RR  / f a k e   1�� "0 theemailaddress theEmailAddress� n  ��STS o  ������  0 emailaddresses emailAddressesT o  ������ 0 
thecontact 
theContact�   email addresses   � �UU     e m a i l   a d d r e s s e s� V��V l ����������  ��  ��  ��  ~   fake loop 0    �WW    f a k e   l o o p   0| m  :;���� y  /fake 0   z �XX  / f a k e   0�� 0 
thecontact 
theContactw l  'Y����Y n   'Z[Z o  #'���� 0 
therecords 
theRecords[ o   #���� 0 thecontacts theContacts��  ��  t  contacts   u �\\  c o n t a c t sr ]^] l     ��������  ��  ��  ^ _`_ l ��a����a I ����bc
�� .sysodlogaskr        TEXTb n ��ded I  ����f���� 0 format_string  f ghg o  ������ 0 prompt_done PROMPT_DONEh i�i J  ��jj klk o  ���~�~ 0 messagessent messagesSentl m�}m o  ���|�| $0 recordsprocessed recordsProcessed�}  �  ��  e o  ���{�{ 0 helper Helperc �zno
�z 
apprn o  ���y�y 0 dialog_title DIALOG_TITLEo �xpq
�x 
btnsp J  ��rr s�ws o  ���v�v 0 	button_ok 	BUTTON_OK�w  q �ut�t
�u 
dfltt o  ���s�s 0 	button_ok 	BUTTON_OK�t  ��  ��  ` uvu l     �r�q�p�r  �q  �p  v wxw l     �oyz�o  y   {   z �{{    {x |}| l     �n~�n  ~ R L 	firstName: Jane, lastName: Doe, displayName: Jane Doe, companyName: Acme,     ��� �   	 f i r s t N a m e :   J a n e ,   l a s t N a m e :   D o e ,   d i s p l a y N a m e :   J a n e   D o e ,   c o m p a n y N a m e :   A c m e ,  } ��� l     �m���m  �  	emailAddresses:   � ���   	 e m a i l A d d r e s s e s :� ��� l     �l���l  �  	{   � ���  	 {� ��� l     �k���k  � 1 +		{theType:Home, theAddress:jane@doe.name},   � ��� V 	 	 { t h e T y p e : H o m e ,   t h e A d d r e s s : j a n e @ d o e . n a m e } ,� ��� l     �j���j  � 4 .		{theType:Work, theAddress:jane.doe@acme.com}   � ��� \ 	 	 { t h e T y p e : W o r k ,   t h e A d d r e s s : j a n e . d o e @ a c m e . c o m }� ��� l     �i���i  �  	}   � ���  	 }� ��� l     �h���h  �  }   � ���  }� ��g� l     �f�e�d�f  �e  �d  �g       �c���c  � �b
�b .aevtoappnull  �   � ****� �a��`�_���^
�a .aevtoappnull  �   � ****� k    ���  W��  `��  g��  n��  u��  ~��  ���  ���  ���  ���  ���  ���  ���  ���  ���  ���  ���  ���  ���  ���  �� &�� .�� _�� ~�� ��� ��� ��� 0�� >�� O�� Z�� c�� l�� z�� ��� ��� ��� ��� ��� ��� ��� �� �� =�� C�� M�� U�� _�� g�� q�� _�]�]  �`  �_  � �\�[�\ 0 
thecontact 
theContact�[ "0 theemailaddress theEmailAddress� � \�Z e�Y l�X s�W z�V ��U ��T ��S ��R ��Q ��P ��O ��N ��M ��L ��K ��J ��I ��H ��G�F�E�D�C�Bh�Ak�@�?�>�=�<�;�:�9�8�7���6���5�4�3�2�1�0�/�.�-�,�+�*�)�(�'�&�%�$�#�"�!� ����������������������
�	��������� �������������Z 0 dialog_title DIALOG_TITLE�Y 0 	button_ok 	BUTTON_OK�X 0 button_cancel BUTTON_CANCEL�W 0 
button_yes 
BUTTON_YES�V 0 	button_no 	BUTTON_NO�U 0 button_mail BUTTON_MAIL�T  0 button_outlook BUTTON_OUTLOOK�S *0 prompt_testing_mode PROMPT_TESTING_MODE�R *0 prompt_what_address PROMPT_WHAT_ADDRESS�Q 0 answer_email ANSWER_EMAIL�P .0 prompt_which_provider PROMPT_WHICH_PROVIDER�O (0 prompt_which_group PROMPT_WHICH_GROUP�N 00 prompt_no_valid_emails PROMPT_NO_VALID_EMAILS�M 60 prompt_which_address_type PROMPT_WHICH_ADDRESS_TYPE�L .0 prompt_which_template PROMPT_WHICH_TEMPLATE�K 0 promt_pausing PROMT_PAUSING�J 0 prompt_done PROMPT_DONE�I &0 reason_no_address REASON_NO_ADDRESS�H 40 reason_duplicate_address REASON_DUPLICATE_ADDRESS�G 60 reason_wrong_address_type REASON_WRONG_ADDRESS_TYPE�F �E 0 thevariables theVariables�D 2�C 0 	batchsize 	batchSize�B 0 thepause thePause
�A 
in D
�@ .sysorpthalis        TEXT
�? .sysoloadscpt        file�> 0 helper Helper
�= 
appr
�< 
btns
�; 
dflt�: 
�9 .sysodlogaskr        TEXT
�8 
rslt
�7 
bhit�6 0 	theplugin 	thePlugin�5 0 testingmode testingMode
�4 
dtxt�3 
�2 
ttxt�1  0 testingaddress testingAddress�0 0 	getgroups 	getGroups�/ 0 
groupnames 
groupNames�. 0 simple_sort  
�- 
prmp�, 
�+ .gtqpchltns    @   @ ns  �* 0 	groupname 	groupName�) 0 getcontacts getContacts�( 0 thecontacts theContacts�' 0 thetypes theTypes
�& .corecnte****       ****
�% 
mlsl�$ 0 gettemplates getTemplates�# 0 thesubjectids theSubjectIds�" 0 	thechoice 	theChoice�! 0 find  �  0 n  � 0 m  
� 
ctxt� 0 theid theID
� .misccurdldt    ��� null� 0 	starttime 	startTime� 0 iterator  � $0 recordsprocessed recordsProcessed� 0 messagessent messagesSent� .0 theaddressesprocessed theAddressesProcessed� 0 
theresults 
theResults� 0 
therecords 
theRecords
� 
kocl
� 
cobj�  0 emailaddresses emailAddresses
� 
leng� 0 displayname displayName� 0 companyname companyName� 0 emailaddress emailAddress
� 
null� 0 type  � 0 sent  �
 
0 reason  �	 � 0 	theresult 	theResult� 0 thetype theType� 0 
theaddress 
theAddress� 0 
getcontent 
getContent� 0 
thecontent 
theContent� 0 personalize  � 0 thename  � 0 therecipient theRecipient�  (0 copyandsendmessage copyAndSendMessage
�� 
bool�� <�� 0 format_string  
�� 
givu�� 
�� .sysodelanull��� ��� nmbr�^��E�O�E�O�E�O�E�O�E�O�E�O�E�O�E�Oa E` Oa E` Oa E` Oa E` Oa E` Oa E` Oa E` Oa E` Oa  E` !Oa "E` #Oa $E` %Oa &E` 'Oa (a )a *a +a ,a -a .a /vE` 0Oa 1E` 2OkE` 3Oa 4a 5a 6l 7j 8E` 9O_ a :�a ;��lva <�a = >O_ ?a @�l  a Aa 5a Bl 7j 8E` COPY *_ ?a @�l  a Da 5a El 7j 8E` COPY hO�a :�a ;��lva <�a = >E` FO_ ?a @�l  fE` FOPY O_ ?a @�l  AeE` FO_ a G_ a :�a ;�kva <�a H >a I,E` JO_ ?f  hY hOPY hO_ Cj+ KE` LO_ 9_ Lk+ ME` LO_ La :�a N_ a O PE` QO_ ?f  hY hO_ C_ Qk+ RE` SO_ ?f  hY hO_ Sa T,j Uj  _ a :�a ;�kva <�a = >OhY hO_ Sa T,a :�a N_ a Vea = PE` TO_ ?f  hY hO_ Cj+ WE` XO_ 9_ Xk+ ME` XO_ Xa :�a N_ a O PE` YO_ ?f  hY hO_ 9a Z_ Yl+ [E` \O_ 9a ]_ Yl+ [E` ^O_ Ya _&[a _\[Z_ \k\Z_ ^k2a _&E` `O*j aE` bOjE` cOjE` dOjE` eOjvE` fOjvE` gO�_ Sa h,[a ia jl Ukh  ikkh_ dkE` dO�a k,a l,j  7a m�a m,a n�a n,a oa pa qa pa rfa s_ #a tE` uOOPY hO�a k,[a ia jl Ukh �kkh_ T�a v, ;a m�a m,a n�a n,a o�a w,a q�a v,a rfa s_ 'a tE` uOOPY H_ f�a w, ;a m�a m,a n�a n,a o�a w,a q�a v,a rfa s_ %a tE` uOOPY hO_ C_ `k+ xE` yO_ 9_ y_ 0��a w,a O+ zE` yOa {a pa wa pa OE` |O_ F a {�a m,a w_ Ja OE` |Y a {�a m,a w�a w,a OE` |O_ C_ `_ y_ |m+ }Oa m�a m,a n�a n,a o�a w,a q�a v,a rea sa pa tE` uO_ f�a w, �a w,_ f6GY hO_ ckE` cO_ ekE` eO_ Sa h,j U_ dj	 _ c_ 2 a ~& JjE` cO_ 9_ *j aa _ 3 a _&kvl+ �a :�a ;�kva �a �a = >Oa _ 3 j �Y hOP[OY�[OY�
OP[OY��[OY��O_ 9_ !_ e_ dlvl+ �a :�a ;�kva <�a = > ascr  ��ޭ