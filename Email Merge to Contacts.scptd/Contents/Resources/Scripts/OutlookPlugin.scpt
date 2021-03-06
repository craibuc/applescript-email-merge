FasdUAS 1.101.10   ��   ��    k             l     ��  ��    W Q --------------------------------------------------------------------------------     � 	 	 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -   
  
 l     ��  ��    : 4 Author:	Craig Buchanan <craig.buchanan@cogniza.com>     �   h   A u t h o r : 	 C r a i g   B u c h a n a n   < c r a i g . b u c h a n a n @ c o g n i z a . c o m >      l     ��  ��      Created:	2013-09-07     �   (   C r e a t e d : 	 2 0 1 3 - 0 9 - 0 7      l     ��  ��    2 , Purpose:	Useful Microsoft Outlook functions     �   X   P u r p o s e : 	 U s e f u l   M i c r o s o f t   O u t l o o k   f u n c t i o n s      l     ��  ��      Installation:      �      I n s t a l l a t i o n :        l     ��   !��     g a			1) Drag script (Contacts Library.applescript) to ~/Library/Scripts/Applications/Outlook folder    ! � " " � 	 	 	 1 )   D r a g   s c r i p t   ( C o n t a c t s   L i b r a r y . a p p l e s c r i p t )   t o   ~ / L i b r a r y / S c r i p t s / A p p l i c a t i o n s / O u t l o o k   f o l d e r   # $ # l     �� % &��   % 9 3			2) Open with AppleScript Editor and compile (?K)    & � ' ' f 	 	 	 2 )   O p e n   w i t h   A p p l e S c r i p t   E d i t o r   a n d   c o m p i l e   (# K ) $  ( ) ( l     �� * +��   * $ 			3) Export as Script (.scpt)    + � , , < 	 	 	 3 )   E x p o r t   a s   S c r i p t   ( . s c p t ) )  - . - l     �� / 0��   / W Q --------------------------------------------------------------------------------    0 � 1 1 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - .  2 3 2 l     ��������  ��  ��   3  4 5 4 l     �� 6 7��   6 W Q --------------------------------------------------------------------------------    7 � 8 8 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 5  9 : 9 l     �� ; <��   ;   unit testing    < � = =    u n i t   t e s t i n g :  > ? > l     �� @ A��   @ W Q --------------------------------------------------------------------------------    A � B B �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ?  C D C l      �� E F��   E��

set theGroups to my getGroups()
(choose from list theGroups with prompt "Which group to use?")
set theChoice to result as text

set theContacts to my getContacts(theChoice)

repeat with theType in (theTypes of theContacts)
	display dialog theType
end repeat

repeat with theContact in (theRecords of theContacts)
	display dialog "FN: " & (firstName of theContact) & "; LN: " & (lastName of theContact) & "; DN: " & (displayName of theContact) & "; CN: " & (companyName of theContact)
	repeat with emailAddress in (emailAddresses of theContact)
		display dialog (theType of emailAddress) & ":" & (theAddress of emailAddress)
	end repeat
end repeat

set theMessage to my getTemplates()
(choose from list theMessage with prompt "Which template to use?")
set theChoice to result as text
display dialog of theChoice

--sendMessage("Outlook - subject", "Message sent by Outlook.", {thename:"First Last", theAddress:"first.last@mailinator.com"})

display dialog (getContent(15002))
    F � G G� 
 
 s e t   t h e G r o u p s   t o   m y   g e t G r o u p s ( ) 
 ( c h o o s e   f r o m   l i s t   t h e G r o u p s   w i t h   p r o m p t   " W h i c h   g r o u p   t o   u s e ? " ) 
 s e t   t h e C h o i c e   t o   r e s u l t   a s   t e x t 
 
 s e t   t h e C o n t a c t s   t o   m y   g e t C o n t a c t s ( t h e C h o i c e ) 
 
 r e p e a t   w i t h   t h e T y p e   i n   ( t h e T y p e s   o f   t h e C o n t a c t s ) 
 	 d i s p l a y   d i a l o g   t h e T y p e 
 e n d   r e p e a t 
 
 r e p e a t   w i t h   t h e C o n t a c t   i n   ( t h e R e c o r d s   o f   t h e C o n t a c t s ) 
 	 d i s p l a y   d i a l o g   " F N :   "   &   ( f i r s t N a m e   o f   t h e C o n t a c t )   &   " ;   L N :   "   &   ( l a s t N a m e   o f   t h e C o n t a c t )   &   " ;   D N :   "   &   ( d i s p l a y N a m e   o f   t h e C o n t a c t )   &   " ;   C N :   "   &   ( c o m p a n y N a m e   o f   t h e C o n t a c t ) 
 	 r e p e a t   w i t h   e m a i l A d d r e s s   i n   ( e m a i l A d d r e s s e s   o f   t h e C o n t a c t ) 
 	 	 d i s p l a y   d i a l o g   ( t h e T y p e   o f   e m a i l A d d r e s s )   &   " : "   &   ( t h e A d d r e s s   o f   e m a i l A d d r e s s ) 
 	 e n d   r e p e a t 
 e n d   r e p e a t 
 
 s e t   t h e M e s s a g e   t o   m y   g e t T e m p l a t e s ( ) 
 ( c h o o s e   f r o m   l i s t   t h e M e s s a g e   w i t h   p r o m p t   " W h i c h   t e m p l a t e   t o   u s e ? " ) 
 s e t   t h e C h o i c e   t o   r e s u l t   a s   t e x t 
 d i s p l a y   d i a l o g   o f   t h e C h o i c e 
 
 - - s e n d M e s s a g e ( " O u t l o o k   -   s u b j e c t " ,   " M e s s a g e   s e n t   b y   O u t l o o k . " ,   { t h e n a m e : " F i r s t   L a s t " ,   t h e A d d r e s s : " f i r s t . l a s t @ m a i l i n a t o r . c o m " } ) 
 
 d i s p l a y   d i a l o g   ( g e t C o n t e n t ( 1 5 0 0 2 ) ) 
 D  H I H l     ��������  ��  ��   I  J K J l     �� L M��   L W Q --------------------------------------------------------------------------------    M � N N �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - K  O P O l     �� Q R��   Q   / unit testing    R � S S    /   u n i t   t e s t i n g P  T U T l     �� V W��   V W Q --------------------------------------------------------------------------------    W � X X �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - U  Y Z Y l     ��������  ��  ��   Z  [ \ [ l     �� ] ^��   ] W Q --------------------------------------------------------------------------------    ^ � _ _ �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - \  ` a ` l     �� b c��   b #  default string localizations    c � d d :   d e f a u l t   s t r i n g   l o c a l i z a t i o n s a  e f e l     �� g h��   g W Q --------------------------------------------------------------------------------    h � i i �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - f  j k j l     ��������  ��  ��   k  l m l l     ��������  ��  ��   m  n o n l     �� p q��   p ( " Purpose: return the plugin's name    q � r r D   P u r p o s e :   r e t u r n   t h e   p l u g i n ' s   n a m e o  s t s l     �� u v��   u   Returns:	Text    v � w w    R e t u r n s : 	 T e x t t  x y x l     ��������  ��  ��   y  z { z i      | } | I      �������� 0 getname getName��  ��   } L      ~ ~ m        � � �  O u t l o o k   P l u g i n {  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � � � Purpose:	Create a new message with the properties (including attachments) of the selected template, altering the Content and the Recipient    � � � �   P u r p o s e : 	 C r e a t e   a   n e w   m e s s a g e   w i t h   t h e   p r o p e r t i e s   ( i n c l u d i n g   a t t a c h m e n t s )   o f   t h e   s e l e c t e d   t e m p l a t e ,   a l t e r i n g   t h e   C o n t e n t   a n d   t h e   R e c i p i e n t �  � � � l     �� � ���   �   Parameters:    � � � �    P a r a m e t e r s : �  � � � l     �� � ���   � . (			theID: the ID of the message template    � � � � P 	 	 	 t h e I D :   t h e   I D   o f   t h e   m e s s a g e   t e m p l a t e �  � � � l     �� � ���   � 2 ,			theContent: the message's content; string    � � � � X 	 	 	 t h e C o n t e n t :   t h e   m e s s a g e ' s   c o n t e n t ;   s t r i n g �  � � � l     �� � ���   � l f			theRecipient: the message's recipient; Record; format: {theName:Jane Doe, theAddress:jane@doe.name}    � � � � � 	 	 	 t h e R e c i p i e n t :   t h e   m e s s a g e ' s   r e c i p i e n t ;   R e c o r d ;   f o r m a t :   { t h e N a m e : J a n e   D o e ,   t h e A d d r e s s : j a n e @ d o e . n a m e } �  � � � l     �� � ���   �   Returns:	nothing    � � � � "   R e t u r n s : 	 n o t h i n g �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� (0 copyandsendmessage copyAndSendMessage �  � � � o      ���� 0 theid theID �  � � � o      ���� 0 
thecontent 
theContent �  ��� � o      ���� 0 therecipient theRecipient��  ��   � k     g � �  � � � l     ��������  ��  ��   �  � � � r      � � � b     	 � � � b      � � � m      � � � � � � < b r / > S e n t   w i t h   < a   h r e f = ' h t t p s : / / g i t h u b . c o m / c r a i b u c / a p p l e s c r i p t - e m a i l - m e r g e ' > A p p l e S c r i p t   E m a i l   M e r g e   ( � n    � � � I    �������� 0 getname getName��  ��   �  f     � m     � � � � � 
 ) < a / > � o      ���� 0 advertisement ADVERTISEMENT �  � � � l   ��������  ��  ��   �  � � � O    e � � � k    d � �  � � � l   ��������  ��  ��   �  � � � l   �� � ���   �  find the template    � � � � " f i n d   t h e   t e m p l a t e �  � � � r     � � � 5    �� ���
�� 
msg  � o    ���� 0 theid theID
�� kfrmID   � o      ���� 0 thetemplate theTemplate �  � � � l   �� � ���   � P J create a new message with the some of the same properties as the template    � � � � �   c r e a t e   a   n e w   m e s s a g e   w i t h   t h e   s o m e   o f   t h e   s a m e   p r o p e r t i e s   a s   t h e   t e m p l a t e �  � � � r    . � � � I   ,���� �
�� .corecrel****      � null��   � �� � �
�� 
kocl � m    ��
�� 
outm � �� ���
�� 
prdt � K    ( � � �� � �
�� 
subj � l     ����� � n      � � � 1     ��
�� 
subj � o    ���� 0 thetemplate theTemplate��  ��   � �� � �
�� 
ctnt � o   ! "���� 0 
thecontent 
theContent � �� ���
�� 
cAtc � l  # & ����� � n   # & � � � 2  $ &��
�� 
cAtc � o   # $���� 0 thetemplate theTemplate��  ��  ��  ��   � o      ���� 0 
themessage 
theMessage �  � � � l  / /�� � ���   �  	 add plug    � � � �    a d d   p l u g �  � � � r   / 8 � � � l  / 4 ����� � b   / 4 � � � l  / 2 ����� � n   / 2 � � � 1   0 2��
�� 
ctnt � o   / 0���� 0 
themessage 
theMessage��  ��   � o   2 3���� 0 advertisement ADVERTISEMENT��  ��   � l      ����� � n       � � � 1   5 7��
�� 
ctnt � o   4 5���� 0 
themessage 
theMessage��  ��   �  � � � l  9 9�� ��      add recipient    �    a d d   r e c i p i e n t �  I  9 \����
�� .corecrel****      � null��   ��
�� 
kocl m   ; <��
�� 
rcpt ��	
�� 
insh o   ? @���� 0 
themessage 
theMessage	 ��
��
�� 
prdt
 K   A X ����
�� 
emad K   D V ��
�� 
pnam l  G L��� n   G L o   H L�~�~ 0 thename   o   G H�}�} 0 therecipient theRecipient��  �   �|�{
�| 
radd l  O T�z�y n   O T o   P T�x�x 0 
theaddress 
theAddress o   O P�w�w 0 therecipient theRecipient�z  �y  �{  ��  ��    l  ] ]�v�v     send the copy    �    s e n d   t h e   c o p y  I  ] b�u�t
�u .mailsendnull���     msg  o   ] ^�s�s 0 
themessage 
theMessage�t   �r l  c c�q�p�o�q  �p  �o  �r   � m                                                                                        OPIM  alis    �  Macintosh HD               ���4H+   
hMicrosoft Outlook.app                                           
s�Țs�        ����  	                Microsoft Office 2011     ���      Ț�F     
h   M  GMacintosh HD:Applications: Microsoft Office 2011: Microsoft Outlook.app   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  8Applications/Microsoft Office 2011/Microsoft Outlook.app  / ��   � !�n! l  f f�m�l�k�m  �l  �k  �n   � "#" l     �j�i�h�j  �i  �h  # $%$ l     �g�f�e�g  �f  �e  % &'& l     �d()�d  ( + % Purpose:	Open a message given its ID   ) �** J   P u r p o s e : 	 O p e n   a   m e s s a g e   g i v e n   i t s   I D' +,+ l     �c-.�c  -   Parameters:   . �//    P a r a m e t e r s :, 010 l     �b23�b  2  			theId: message ID   3 �44 ( 	 	 	 t h e I d :   m e s s a g e   I D1 565 l     �a78�a  7   Returns:	a Message   8 �99 &   R e t u r n s : 	 a   M e s s a g e6 :;: l     �`�_�^�`  �_  �^  ; <=< i    >?> I      �]@�\�] 0 
getcontent 
getContent@ A�[A o      �Z�Z 0 theid theID�[  �\  ? k     BB CDC l     �Y�X�W�Y  �X  �W  D EFE O     GHG L    II c    JKJ l   L�V�UL l   M�T�SM n    NON 1   	 �R
�R 
ctntO 5    	�QP�P
�Q 
msg P o    �O�O 0 theid theID
�P kfrmID  �T  �S  �V  �U  K m    �N
�N 
ctxtH m     QQ                                                                                  OPIM  alis    �  Macintosh HD               ���4H+   
hMicrosoft Outlook.app                                           
s�Țs�        ����  	                Microsoft Office 2011     ���      Ț�F     
h   M  GMacintosh HD:Applications: Microsoft Office 2011: Microsoft Outlook.app   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  8Applications/Microsoft Office 2011/Microsoft Outlook.app  / ��  F R�MR l   �L�K�J�L  �K  �J  �M  = STS l     �I�H�G�I  �H  �G  T UVU l     �F�E�D�F  �E  �D  V WXW l     �CYZ�C  Y P J Purpose:	Assemble a list of subjects of each message in the Drafts folder   Z �[[ �   P u r p o s e : 	 A s s e m b l e   a   l i s t   o f   s u b j e c t s   o f   e a c h   m e s s a g e   i n   t h e   D r a f t s   f o l d e rX \]\ l     �B^_�B  ^ X R Returns:	List of Message subjects and IDs. The ID can be used to open the message   _ �`` �   R e t u r n s : 	 L i s t   o f   M e s s a g e   s u b j e c t s   a n d   I D s .   T h e   I D   c a n   b e   u s e d   t o   o p e n   t h e   m e s s a g e] aba l     �A�@�?�A  �@  �?  b cdc i    efe I      �>�=�<�> 0 gettemplates getTemplates�=  �<  f k     Dgg hih l     �;�:�9�;  �:  �9  i jkj O     Blml k    Ann opo l   �8�7�6�8  �7  �6  p qrq l   �5st�5  s 7 1 generate a List of drafts in the 'Drafts' folder   t �uu b   g e n e r a t e   a   L i s t   o f   d r a f t s   i n   t h e   ' D r a f t s '   f o l d e rr vwv r    xyx J    �4�4  y o      �3�3 0 thesubjectids theSubjectIdsw z{z l  	 	�2�1�0�2  �1  �0  { |}| r   	 ~~ n  	 ��� 2    �/
�/ 
msg � l  	 ��.�-� 1   	 �,
�, 
pDrt�.  �-   o      �+�+ 0 themessages theMessages} ��� X    <��*�� k   ! 7�� ��� l  ! !�)���)  �   Subject [ID]   � ���    S u b j e c t   [ I D ]� ��� r   ! 2��� b   ! 0��� b   ! .��� b   ! (��� l  ! &��(�'� c   ! &��� n   ! $��� 1   " $�&
�& 
subj� o   ! "�%�% 0 
themessage 
theMessage� m   $ %�$
�$ 
ctxt�(  �'  � m   & '�� ���    [� l  ( -��#�"� c   ( -��� n   ( +��� 1   ) +�!
�! 
ID  � o   ( )� �  0 
themessage 
theMessage� m   + ,�
� 
ctxt�#  �"  � m   . /�� ���  ]� o      �� 0 thesubjectid theSubjectId� ��� s   3 7��� o   3 4�� 0 thesubjectid theSubjectId� n      ���  ;   5 6� o   4 5�� 0 thesubjectids theSubjectIds�  �* 0 
themessage 
theMessage� o    �� 0 themessages theMessages� ��� l  = =����  �  �  � ��� L   = ?�� o   = >�� 0 thesubjectids theSubjectIds� ��� l  @ @����  �  �  �  m m     ��                                                                                  OPIM  alis    �  Macintosh HD               ���4H+   
hMicrosoft Outlook.app                                           
s�Țs�        ����  	                Microsoft Office 2011     ���      Ț�F     
h   M  GMacintosh HD:Applications: Microsoft Office 2011: Microsoft Outlook.app   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  8Applications/Microsoft Office 2011/Microsoft Outlook.app  / ��  k ��� l  C C����  �  �  �  d ��� l     ����  �  �  � ��� l     �
�	��
  �	  �  � ��� l     ����  � > 8 Purpose:	Convert selected Contacts to a List of Records   � ��� p   P u r p o s e : 	 C o n v e r t   s e l e c t e d   C o n t a c t s   t o   a   L i s t   o f   R e c o r d s� ��� l     ����  �   Parameters:   � ���    P a r a m e t e r s :� ��� l     ����  � ` Z			* categoryName - unused; matches interface of similar function in Contacts Library.scpt   � ��� � 	 	 	 *   c a t e g o r y N a m e   -   u n u s e d ;   m a t c h e s   i n t e r f a c e   o f   s i m i l a r   f u n c t i o n   i n   C o n t a c t s   L i b r a r y . s c p t� ��� l     ����  � y s Returns:	A Record that contains the distinct list of address types (e.g. 'work') and a List (unsorted) of Records:   � ��� �   R e t u r n s : 	 A   R e c o r d   t h a t   c o n t a i n s   t h e   d i s t i n c t   l i s t   o f   a d d r e s s   t y p e s   ( e . g .   ' w o r k ' )   a n d   a   L i s t   ( u n s o r t e d )   o f   R e c o r d s :� ��� l     ����  �  	{   � ���  	 {� ��� l     ����  � ! 		theTypes: {'home','work'}   � ��� 6 	 	 t h e T y p e s :   { ' h o m e ' , ' w o r k ' }� ��� l     ����  �  		theRecords:{   � ���  	 	 t h e R e c o r d s : {� ��� l     � ���   � � � 			{ firstName: Jane, lastName: Doe, displayName: Jane Doe, companyName: Acme, emailAddresses: {{theType:"home", theAddress: jane@doe.name},{theType:"work", theAddress: jane.doe@acme.com}} },   � ����   	 	 	 {   f i r s t N a m e :   J a n e ,   l a s t N a m e :   D o e ,   d i s p l a y N a m e :   J a n e   D o e ,   c o m p a n y N a m e :   A c m e ,   e m a i l A d d r e s s e s :   { { t h e T y p e : " h o m e " ,   t h e A d d r e s s :   j a n e @ d o e . n a m e } , { t h e T y p e : " w o r k " ,   t h e A d d r e s s :   j a n e . d o e @ a c m e . c o m } }   } ,� ��� l     ������  � 
 			�   � ���  	 	 	 &� ��� l     ������  � � � 			{ firstName: John, lastName: Doe, displayName: John Doe, companyName: Acme, emailAddresses: {{theType:"home", theAddress: john@doe.name},{theType:"work", theAddress: john.doe@acme.com}} }   � ���~   	 	 	 {   f i r s t N a m e :   J o h n ,   l a s t N a m e :   D o e ,   d i s p l a y N a m e :   J o h n   D o e ,   c o m p a n y N a m e :   A c m e ,   e m a i l A d d r e s s e s :   { { t h e T y p e : " h o m e " ,   t h e A d d r e s s :   j o h n @ d o e . n a m e } , { t h e T y p e : " w o r k " ,   t h e A d d r e s s :   j o h n . d o e @ a c m e . c o m } }   }� ��� l     ������  � 
 			}   � ���  	 	 	 }� ��� l     ������  �  	}   � ���  	 }� ��� l     ������  �   TODOs:	   � ���    T O D O s : 	� ��� l     ������  � * $			* Filter Contacts by categoryName   � ��� H 	 	 	 *   F i l t e r   C o n t a c t s   b y   c a t e g o r y N a m e� ��� l     ������  � > 8			* Identify primary email address; not exposed by API?   � �   p 	 	 	 *   I d e n t i f y   p r i m a r y   e m a i l   a d d r e s s ;   n o t   e x p o s e d   b y   A P I ?�  l     ��������  ��  ��    i     I      ������ 0 getcontacts getContacts �� o      ���� 0 categoryname categoryName��  ��   k    \		 

 l     ��������  ��  ��    O    Z k   Y  l   ��������  ��  ��    r     J    ����   o      ���� 0 thecontacts theContacts  l  	 	��������  ��  ��    Z   	 #�� =  	  l  	  ����  c   	 !"! o   	 
���� 0 categoryname categoryName" m   
 ��
�� 
ctxt��  ��   l   #����# c    $%$ m    && �'' ( < < S e l e c t e d   O b j e c t s > >% m    ��
�� 
ctxt��  ��   r    ()( 1    ��
�� 
SeOb) o      ���� 0 thecontacts theContacts��   k    #** +,+ I    ��-��
�� .sysodlogaskr        TEXT- m    .. �// X S e a r c h   b y   C a t e g o r y   h a s   n o t   b e e n   i m p l e m e n t e d .��  , 010 l  ! !��23��  2 0 * TODO: add logic to filter by categoryName   3 �44 T   T O D O :   a d d   l o g i c   t o   f i l t e r   b y   c a t e g o r y N a m e1 5��5 L   ! #66 m   ! "��
�� boovfals��   787 l  $ $��������  ��  ��  8 9:9 l  $ $��;<��  ; 8 2 if there are no contacts selected, warn then quit   < �== d   i f   t h e r e   a r e   n o   c o n t a c t s   s e l e c t e d ,   w a r n   t h e n   q u i t: >?> Z   $ ;@A����@ =  $ (BCB o   $ %���� 0 thecontacts theContactsC J   % '����  A k   + 7DD EFE I  + 4��GH
�� .sysodlogaskr        TEXTG m   + ,II �JJ � N o   c o n t a c t s   h a v e   b e e n   s e l e c t e d .     P l e a s e   s e l e c t   o n e   o r   m o r e   c o n t a c t s   i n   t h e   C o n t a c t s   p a n e   (# 3 ) .H ��KL
�� 
apprK m   - .MM �NN  O u t l o o k   L i b r a r yL ��O��
�� 
dispO m   / 0���� ��  F P��P L   5 7QQ m   5 6��
�� boovfals��  ��  ��  ? RSR l  < <��������  ��  ��  S TUT l  < <��VW��  V B <test type of selected items to ensure that they are Contacts   W �XX x t e s t   t y p e   o f   s e l e c t e d   i t e m s   t o   e n s u r e   t h a t   t h e y   a r e   C o n t a c t sU YZY Z   < W[\����[ >  < D]^] n   < B_`_ 1   @ B��
�� 
pcls` n   < @aba 4   = @��c
�� 
cobjc m   > ?���� b o   < =���� 0 thecontacts theContacts^ m   B C��
�� 
cAbE\ k   G Sdd efe I  G P��gh
�� .sysodlogaskr        TEXTg m   G Hii �jj � T h e   i t e m s   t h a t   h a v e   b e e n   s e l e c t e d   a r e   N O T   c o n t a c t s .     P l e a s e   s e l e c t   o n e   o r   m o r e   c o n t a c t s   i n   t h e   C o n t a c t s   p a n e   (# 3 ) .h ��kl
�� 
apprk m   I Jmm �nn  O u t l o o k   L i b r a r yl ��o��
�� 
dispo m   K L���� ��  f p��p L   Q Sqq m   Q R��
�� boovfals��  ��  ��  Z rsr l  X X��������  ��  ��  s tut l  X X��vw��  v 0 *unique list of address types (e.g. 'work')   w �xx T u n i q u e   l i s t   o f   a d d r e s s   t y p e s   ( e . g .   ' w o r k ' )u yzy r   X \{|{ J   X Z����  | o      ���� 0 	thetypesx 	theTypesXz }~} r   ] a� J   ] _����  � o      ���� 0 therecordsx theRecordsX~ ��� l  b b��������  ��  ��  � ��� X   bK����� k   tF�� ��� l  t t��������  ��  ��  � ��� l  t t������  � : 4 store relevant properties in Record for easy access   � ��� h   s t o r e   r e l e v a n t   p r o p e r t i e s   i n   R e c o r d   f o r   e a s y   a c c e s s� ��� r   t ���� K   t ��� ������ 0 	firstname 	firstName� m   w z��
�� 
null� ������ 0 lastname lastName� m   } ���
�� 
null� ������ 0 displayname displayName� m   � ���
�� 
null� ������ 0 companyname companyName� m   � ���
�� 
null� �������  0 emailaddresses emailAddresses� m   � ���
�� 
null��  � o      ���� 0 	therecord 	theRecord� ��� l  � ���������  ��  ��  � ��� r   � ���� n   � ���� 1   � ���
�� 
pFrN� o   � ����� 0 
thecontact 
theContact� n      ��� o   � ����� 0 	firstname 	firstName� o   � ����� 0 	therecord 	theRecord� ��� r   � ���� n   � ���� 1   � ���
�� 
pLsN� o   � ����� 0 
thecontact 
theContact� n      ��� o   � ����� 0 lastname lastName� o   � ����� 0 	therecord 	theRecord� ��� r   � ���� n   � ���� 1   � ���
�� 
dspn� o   � ����� 0 
thecontact 
theContact� n      ��� o   � ����� 0 displayname displayName� o   � ����� 0 	therecord 	theRecord� ��� r   � ���� n   � ���� 1   � ���
�� 
Cmpy� o   � ����� 0 
thecontact 
theContact� n      ��� o   � ����� 0 companyname companyName� o   � ����� 0 	therecord 	theRecord� ��� l  � ���������  ��  ��  � ��� r   � ���� J   � �����  � o      ���� 0 addresslist addressList� ��� r   � ���� J   � �����  � o      ����  0 emailaddresses emailAddresses� ��� l  � �����~��  �  �~  � ��� r   � ���� n   � ���� 1   � ��}
�} 
EmAd� o   � ��|�| 0 
thecontact 
theContact� o      �{�{ 0 theaddresses theAddresses� ��� X   �7��z�� k   �2�� ��� l  � ��y���y  � # 				display dialog theAddress   � ��� : 	 	 	 	 d i s p l a y   d i a l o g   t h e A d d r e s s� ��� l  � ��x���x  � / ) build unique type list (for theContacts)   � ��� R   b u i l d   u n i q u e   t y p e   l i s t   ( f o r   t h e C o n t a c t s )� ��� Z   ����w�v� H   � ��� E   � ���� o   � ��u�u 0 	thetypesx 	theTypesX� l  � ���t�s� c   � ���� n   � ���� m   � ��r
�r 
type� o   � ��q�q 0 
theaddress 
theAddress� m   � ��p
�p 
ctxt�t  �s  � k   ��� ��� l  � ��o���o  � ? 9set theTypesX to theTypesX & (type of theAddress as text)   � ��� r s e t   t h e T y p e s X   t o   t h e T y p e s X   &   ( t y p e   o f   t h e A d d r e s s   a s   t e x t )� ��� l  � ��n���n  � 0 *copy emailAddress to end of emailAddresses   � ��� T c o p y   e m a i l A d d r e s s   t o   e n d   o f   e m a i l A d d r e s s e s� ��m� s   ���� l  ���l�k� c   ���� n   � ���� m   � ��j
�j 
type� o   � ��i�i 0 
theaddress 
theAddress� m   � �h
�h 
ctxt�l  �k  � n      � �  ;    o  �g�g 0 	thetypesx 	theTypesX�m  �w  �v  �  l 		�f�e�d�f  �e  �d    l 		�c�c   0 *build unique address list (for theContact)    � T b u i l d   u n i q u e   a d d r e s s   l i s t   ( f o r   t h e C o n t a c t ) 	 l 		�b
�b  
 B <if addressList does not contain (address of theAddress) then    � x i f   a d d r e s s L i s t   d o e s   n o t   c o n t a i n   ( a d d r e s s   o f   t h e A d d r e s s )   t h e n	  l 		�a�`�_�a  �`  �_    l 		�^�^   3 - {theType:Home, theAddress:jane@doe.name}				    � Z   { t h e T y p e : H o m e ,   t h e A d d r e s s : j a n e @ d o e . n a m e } 	 	 	 	  r  	! K  	 �]�] 0 thetype theType l �\�[ c   n   m  �Z
�Z 
type o  �Y�Y 0 
theaddress 
theAddress m  �X
�X 
ctxt�\  �[   �W �V�W 0 
theaddress 
theAddress  l !�U�T! c  "#" n  $%$ 1  �S
�S 
radd% o  �R�R 0 
theaddress 
theAddress# m  �Q
�Q 
ctxt�U  �T  �V   o      �P�P 0 emailaddress emailAddress &'& s  "&()( o  "#�O�O 0 emailaddress emailAddress) n      *+*  ;  $%+ o  #$�N�N  0 emailaddresses emailAddresses' ,-, l ''�M�L�K�M  �L  �K  - ./. r  '0010 b  '.232 o  '(�J�J 0 addresslist addressList3 l (-4�I�H4 n  (-565 1  )-�G
�G 
radd6 o  ()�F�F 0 
theaddress 
theAddress�I  �H  1 o      �E�E 0 addresslist addressList/ 787 l 11�D�C�B�D  �C  �B  8 9:9 l 11�A;<�A  ;  end if   < �==  e n d   i f: >�@> l 11�?�>�=�?  �>  �=  �@  �z 0 
theaddress 
theAddress� l  � �?�<�;? o   � ��:�: 0 theaddresses theAddresses�<  �;  � @A@ l 88�9�8�7�9  �8  �7  A BCB r  8?DED o  89�6�6  0 emailaddresses emailAddressesE n      FGF o  :>�5�5  0 emailaddresses emailAddressesG o  9:�4�4 0 	therecord 	theRecordC HIH s  @DJKJ o  @A�3�3 0 	therecord 	theRecordK n      LML  ;  BCM o  AB�2�2 0 therecordsx theRecordsXI N�1N l EE�0�/�.�0  �/  �.  �1  �� 0 
thecontact 
theContact� o   e f�-�- 0 thecontacts theContacts� OPO l LL�,�+�*�,  �+  �*  P QRQ l LL�)ST�)  S  		return theRecords   T �UU & 	 	 r e t u r n   t h e R e c o r d sR VWV L  LWXX K  LVYY �(Z[�( 0 thetypes theTypesZ o  OP�'�' 0 	thetypesx 	theTypesX[ �&\�%�& 0 
therecords 
theRecords\ o  ST�$�$ 0 therecordsx theRecordsX�%  W ]�#] l XX�"�!� �"  �!  �   �#   m     ^^                                                                                  OPIM  alis    �  Macintosh HD               ���4H+   
hMicrosoft Outlook.app                                           
s�Țs�        ����  	                Microsoft Office 2011     ���      Ț�F     
h   M  GMacintosh HD:Applications: Microsoft Office 2011: Microsoft Outlook.app   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  8Applications/Microsoft Office 2011/Microsoft Outlook.app  / ��   _�_ l [[����  �  �  �   `a` l     ����  �  �  a bcb l     ����  �  �  c ded l     �fg�  f � { Purpose:	Collect names of Categories; name of function matches equivalent in OS X Contacts Library.scpt (pseudo-interface)   g �hh �   P u r p o s e : 	 C o l l e c t   n a m e s   o f   C a t e g o r i e s ;   n a m e   o f   f u n c t i o n   m a t c h e s   e q u i v a l e n t   i n   O S   X   C o n t a c t s   L i b r a r y . s c p t   ( p s e u d o - i n t e r f a c e )e iji l     �kl�  k - ' Returns:	List (unsorted) of Categories   l �mm N   R e t u r n s : 	 L i s t   ( u n s o r t e d )   o f   C a t e g o r i e sj non l     �pq�  p � � Notes:		"<<Selected Objects>>" allows any artibratry list (e.g. the contents of a Smart Folder) of contacts to be used as a data source   q �rr   N o t e s : 	 	 " < < S e l e c t e d   O b j e c t s > > "   a l l o w s   a n y   a r t i b r a t r y   l i s t   ( e . g .   t h e   c o n t e n t s   o f   a   S m a r t   F o l d e r )   o f   c o n t a c t s   t o   b e   u s e d   a s   a   d a t a   s o u r c eo sts l     ����  �  �  t u�u i    vwv I      ���� 0 	getgroups 	getGroups�  �  w k     xx yzy l     ��
�	�  �
  �	  z {|{ O     }~} k     ��� l   ����  �  �  � ��� l   ����  � $ generate a List of group names   � ��� < g e n e r a t e   a   L i s t   o f   g r o u p   n a m e s� ��� r    	��� J    �� ��� m    �� ��� ( < < S e l e c t e d   O b j e c t s > >�  � o      �� 0 
groupnames 
groupNames� ��� l  
 
��� �  �  �   � ��� l  
 
������  � . (	repeat with groupItem in every category   � ��� P 	 r e p e a t   w i t h   g r o u p I t e m   i n   e v e r y   c a t e g o r y� ��� l  
 
������  � = 7		copy (name of groupItem) as text to end of groupNames   � ��� n 	 	 c o p y   ( n a m e   o f   g r o u p I t e m )   a s   t e x t   t o   e n d   o f   g r o u p N a m e s� ��� l  
 
������  �  	end repeat   � ���  	 e n d   r e p e a t� ��� l  
 
��������  ��  ��  � ��� L   
 �� o   
 ���� 0 
groupnames 
groupNames� ���� l   ��������  ��  ��  ��  ~ m     ��                                                                                  OPIM  alis    �  Macintosh HD               ���4H+   
hMicrosoft Outlook.app                                           
s�Țs�        ����  	                Microsoft Office 2011     ���      Ț�F     
h   M  GMacintosh HD:Applications: Microsoft Office 2011: Microsoft Outlook.app   ,  M i c r o s o f t   O u t l o o k . a p p    M a c i n t o s h   H D  8Applications/Microsoft Office 2011/Microsoft Outlook.app  / ��  | ���� l   ��������  ��  ��  ��  �       �����������  � �������������� 0 getname getName�� (0 copyandsendmessage copyAndSendMessage�� 0 
getcontent 
getContent�� 0 gettemplates getTemplates�� 0 getcontacts getContacts�� 0 	getgroups 	getGroups� �� }���������� 0 getname getName��  ��  �  �  �� �� �� ����������� (0 copyandsendmessage copyAndSendMessage�� ����� �  �������� 0 theid theID�� 0 
thecontent 
theContent�� 0 therecipient theRecipient��  � �������������� 0 theid theID�� 0 
thecontent 
theContent�� 0 therecipient theRecipient�� 0 advertisement ADVERTISEMENT�� 0 thetemplate theTemplate�� 0 
themessage 
theMessage�  ��� � ���������������������������������������� 0 getname getName
�� 
msg 
�� kfrmID  
�� 
kocl
�� 
outm
�� 
prdt
�� 
subj
�� 
ctnt
�� 
cAtc�� �� 
�� .corecrel****      � null
�� 
rcpt
�� 
insh
�� 
emad
�� 
pnam�� 0 thename  
�� 
radd�� 0 
theaddress 
theAddress
�� .mailsendnull���     msg �� h�)j+ %�%E�O� V*��0E�O*�����,���-�� E�O��,�%��,FO*��a ��a a �a ,a �a ,�l� O�j OPUOP� ��?���������� 0 
getcontent 
getContent�� ����� �  ���� 0 theid theID��  � ���� 0 theid theID� Q��������
�� 
msg 
�� kfrmID  
�� 
ctnt
�� 
ctxt�� � *��0�,�&UOP� ��f���������� 0 gettemplates getTemplates��  ��  � ���������� 0 thesubjectids theSubjectIds�� 0 themessages theMessages�� 0 
themessage 
theMessage�� 0 thesubjectid theSubjectId� �������������������
�� 
pDrt
�� 
msg 
�� 
kocl
�� 
cobj
�� .corecnte****       ****
�� 
subj
�� 
ctxt
�� 
ID  �� E� ?jvE�O*�,�-E�O *�[��l kh ��,�&�%��,�&%�%E�O��6G[OY��O�OPUOP� ������������ 0 getcontacts getContacts�� ����� �  ���� 0 categoryname categoryName��  � ������������������������ 0 categoryname categoryName�� 0 thecontacts theContacts�� 0 	thetypesx 	theTypesX�� 0 therecordsx theRecordsX�� 0 
thecontact 
theContact�� 0 	therecord 	theRecord�� 0 addresslist addressList��  0 emailaddresses emailAddresses�� 0 theaddresses theAddresses�� 0 
theaddress 
theAddress�� 0 emailaddress emailAddress� $^��&��.��I��M����������im��������������������������������~�}�|�{
�� 
ctxt
�� 
SeOb
�� .sysodlogaskr        TEXT
�� 
appr
�� 
disp�� 
�� 
cobj
�� 
pcls
�� 
cAbE
�� 
kocl
�� .corecnte****       ****�� 0 	firstname 	firstName
�� 
null�� 0 lastname lastName�� 0 displayname displayName�� 0 companyname companyName��  0 emailaddresses emailAddresses�� 

�� 
pFrN
�� 
pLsN
�� 
dspn
�� 
Cmpy
�� 
EmAd
�� 
type� 0 thetype theType�~ 0 
theaddress 
theAddress
�} 
radd�| 0 thetypes theTypes�{ 0 
therecords 
theRecords��]�WjvE�O��&��&  
*�,E�Y 
�j OfO�jv  ����k� OfY hO��k/�,� ����k� OfY hOjvE�OjvE�O �[a �l kh a a a a a a a a a a a E�O�a ,�a ,FO�a ,�a ,FO�a ,�a ,FO�a ,�a ,FOjvE�OjvE�O�a ,E�O [�[a �l kh 	��a ,�& �a ,�&�6GY hOa �a ,�&a  �a !,�&�E�O��6GO��a !,%E�OP[OY��O��a ,FO��6GOP[OY�(Oa "�a #��OPUOP� �zw�y�x���w�z 0 	getgroups 	getGroups�y  �x  � �v�v 0 
groupnames 
groupNames� ���w � �kvE�O�OPUOPascr  ��ޭ