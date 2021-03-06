FasdUAS 1.101.10   ��   ��    k             l     ��  ��    W Q --------------------------------------------------------------------------------     � 	 	 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -   
  
 l     ��  ��    : 4 Author:	Craig Buchanan <craig.buchanan@cogniza.com>     �   h   A u t h o r : 	 C r a i g   B u c h a n a n   < c r a i g . b u c h a n a n @ c o g n i z a . c o m >      l     ��  ��      Created:	2013-09-07     �   (   C r e a t e d : 	 2 0 1 3 - 0 9 - 0 7      l     ��  ��    B < Purpose:	Useful OS X Mail/Address Book (Contacts) functions     �   x   P u r p o s e : 	 U s e f u l   O S   X   M a i l / A d d r e s s   B o o k   ( C o n t a c t s )   f u n c t i o n s      l     ��  ��      Installation:      �      I n s t a l l a t i o n :        l     ��   !��     d ^			1) Drag script (Contacts Library.applescript) to ~/Library/Scripts/Applications/Mail folder    ! � " " � 	 	 	 1 )   D r a g   s c r i p t   ( C o n t a c t s   L i b r a r y . a p p l e s c r i p t )   t o   ~ / L i b r a r y / S c r i p t s / A p p l i c a t i o n s / M a i l   f o l d e r   # $ # l     �� % &��   % 9 3			2) Open with AppleScript Editor and compile (?K)    & � ' ' f 	 	 	 2 )   O p e n   w i t h   A p p l e S c r i p t   E d i t o r   a n d   c o m p i l e   (# K ) $  ( ) ( l     �� * +��   * $ 			3) Export as Script (.scpt)    + � , , < 	 	 	 3 )   E x p o r t   a s   S c r i p t   ( . s c p t ) )  - . - l     �� / 0��   / W Q --------------------------------------------------------------------------------    0 � 1 1 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - .  2 3 2 l     ��������  ��  ��   3  4 5 4 l     �� 6 7��   6 W Q --------------------------------------------------------------------------------    7 � 8 8 �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 5  9 : 9 l     �� ; <��   ;   unit testing    < � = =    u n i t   t e s t i n g :  > ? > l     �� @ A��   @ W Q --------------------------------------------------------------------------------    A � B B �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ?  C D C l      �� E F��   ED>

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
		display dialog (theType of emailAddress as text) & ":" & (theAddress of emailAddress as text)
	end repeat
end repeat

set theTemplates to my getTemplates()
(choose from list theTemplates with prompt "Which template to use?")
set theChoice to result as text
display dialog theChoice
    F � G G| 
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
 	 	 d i s p l a y   d i a l o g   ( t h e T y p e   o f   e m a i l A d d r e s s   a s   t e x t )   &   " : "   &   ( t h e A d d r e s s   o f   e m a i l A d d r e s s   a s   t e x t ) 
 	 e n d   r e p e a t 
 e n d   r e p e a t 
 
 s e t   t h e T e m p l a t e s   t o   m y   g e t T e m p l a t e s ( ) 
 ( c h o o s e   f r o m   l i s t   t h e T e m p l a t e s   w i t h   p r o m p t   " W h i c h   t e m p l a t e   t o   u s e ? " ) 
 s e t   t h e C h o i c e   t o   r e s u l t   a s   t e x t 
 d i s p l a y   d i a l o g   t h e C h o i c e 
 D  H I H l     ��������  ��  ��   I  J K J l     �� L M��   L * $set theContent to getContent(100940)    M � N N H s e t   t h e C o n t e n t   t o   g e t C o n t e n t ( 1 0 0 9 4 0 ) K  O P O l     �� Q R��   Q  display dialog theContent    R � S S 2 d i s p l a y   d i a l o g   t h e C o n t e n t P  T U T l     ��������  ��  ��   U  V W V l     �� X Y��   X } wsendMessage("Mail - subject", "Message sent by Mail.", {theName:"First Last", theAddress:"email.merge@mailinator.com"})    Y � Z Z � s e n d M e s s a g e ( " M a i l   -   s u b j e c t " ,   " M e s s a g e   s e n t   b y   M a i l . " ,   { t h e N a m e : " F i r s t   L a s t " ,   t h e A d d r e s s : " e m a i l . m e r g e @ m a i l i n a t o r . c o m " } ) W  [ \ [ l     ��������  ��  ��   \  ] ^ ] l     �� _ `��   _ W Q --------------------------------------------------------------------------------    ` � a a �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ^  b c b l     �� d e��   d   / unit testing    e � f f    /   u n i t   t e s t i n g c  g h g l     �� i j��   i W Q --------------------------------------------------------------------------------    j � k k �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - h  l m l l     ��������  ��  ��   m  n o n l     �� p q��   p W Q --------------------------------------------------------------------------------    q � r r �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - o  s t s l     �� u v��   u #  default string localizations    v � w w :   d e f a u l t   s t r i n g   l o c a l i z a t i o n s t  x y x l     �� z {��   z W Q --------------------------------------------------------------------------------    { � | | �   - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - y  } ~ } l     ��������  ��  ��   ~   �  l     ��������  ��  ��   �  � � � l     �� � ���   � ( " Purpose: return the plugin's name    � � � � D   P u r p o s e :   r e t u r n   t h e   p l u g i n ' s   n a m e �  � � � l     �� � ���   �   Returns:	Text    � � � �    R e t u r n s : 	 T e x t �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � i      � � � I      �������� 0 getname getName��  ��   � L      � � m      � � � � � : O S   X   M a i l / A d d r e s s   B o o k   P l u g i n �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � � � Purpose:	Create a new message with the properties (including attachments) of the selected template, altering the Content and the Recipient    � � � �   P u r p o s e : 	 C r e a t e   a   n e w   m e s s a g e   w i t h   t h e   p r o p e r t i e s   ( i n c l u d i n g   a t t a c h m e n t s )   o f   t h e   s e l e c t e d   t e m p l a t e ,   a l t e r i n g   t h e   C o n t e n t   a n d   t h e   R e c i p i e n t �  � � � l     �� � ���   �   Parameters:    � � � �    P a r a m e t e r s : �  � � � l     �� � ���   � . (			theID: the ID of the message template    � � � � P 	 	 	 t h e I D :   t h e   I D   o f   t h e   m e s s a g e   t e m p l a t e �  � � � l     �� � ���   � 2 ,			theContent: the message's content; string    � � � � X 	 	 	 t h e C o n t e n t :   t h e   m e s s a g e ' s   c o n t e n t ;   s t r i n g �  � � � l     �� � ���   � l f			theRecipient: the message's recipient; Record; format: {theName:Jane Doe, theAddress:jane@doe.name}    � � � � � 	 	 	 t h e R e c i p i e n t :   t h e   m e s s a g e ' s   r e c i p i e n t ;   R e c o r d ;   f o r m a t :   { t h e N a m e : J a n e   D o e ,   t h e A d d r e s s : j a n e @ d o e . n a m e } �  � � � l     �� � ���   �   Returns:	nothing    � � � � "   R e t u r n s : 	 n o t h i n g �  � � � l     ��������  ��  ��   �  � � � i     � � � I      �� ����� (0 copyandsendmessage copyAndSendMessage �  � � � o      ���� 0 theid theID �  � � � o      ���� 0 
thecontent 
theContent �  ��� � o      ���� 0 therecipient theRecipient��  ��   � k     x � �  � � � l     ��������  ��  ��   �  � � � r      � � � b     	 � � � b      � � � m      � � � � � � < b r / > S e n t   w i t h   < a   h r e f = ' h t t p s : / / g i t h u b . c o m / c r a i b u c / a p p l e s c r i p t - e m a i l - m e r g e ' > A p p l e S c r i p t   E m a i l   M e r g e   ( � n    � � � I    �������� 0 getname getName��  ��   �  f     � m     � � � � �  )   < a / > � o      ���� 0 advertisement ADVERTISEMENT �  � � � l   ��������  ��  ��   �  � � � O    v � � � k    u � �  � � � l   ��������  ��  ��   �  � � � l   �� � ���   �  find the template    � � � � " f i n d   t h e   t e m p l a t e �  � � � l   �� � ���   � N Hreturn the content of (first message of drafts mailbox whose id = theID)    � � � � � r e t u r n   t h e   c o n t e n t   o f   ( f i r s t   m e s s a g e   o f   d r a f t s   m a i l b o x   w h o s e   i d   =   t h e I D ) �  � � � l   �� � ���   � &  set theTemplate to message theID    � � � � @ s e t   t h e T e m p l a t e   t o   m e s s a g e   t h e I D �  � � � r    ! � � � l    ����� � l    ����� � 6    � � � n     � � � 4   �� �
�� 
mssg � m    ����  � 1    ��
�� 
drmb � =     � � � 1    ��
�� 
ID   � o    ���� 0 theid theID��  ��  ��  ��   � o      ���� 0 thetemplate theTemplate �  � � � l  " "�� � ���   � P J create a new message with the some of the same properties as the template    � � � � �   c r e a t e   a   n e w   m e s s a g e   w i t h   t h e   s o m e   o f   t h e   s a m e   p r o p e r t i e s   a s   t h e   t e m p l a t e �  �  � r   " 8 I  " 6����
�� .corecrel****      � null��   ��
�� 
kocl m   $ %��
�� 
bcke ����
�� 
prdt K   & 2 ��	
�� 
subj l  ' *
����
 n   ' * 1   ( *��
�� 
subj o   ' (���� 0 thetemplate theTemplate��  ��  	 ��
�� 
ctnt o   + ,���� 0 
thecontent 
theContent ����
�� 
atts l  - 0���� n   - 0 2  . 0��
�� 
atts o   - .���� 0 thetemplate theTemplate��  ��  ��  ��   o      ���� 0 
themessage 
theMessage   l  9 9����    	 add plug    �    a d d   p l u g  r   9 B l  9 >���� b   9 > l  9 <���� n   9 < !  1   : <�
� 
ctnt! o   9 :�~�~ 0 
themessage 
theMessage��  ��   o   < =�}�} 0 advertisement ADVERTISEMENT��  ��   l     "�|�{" n      #$# 1   ? A�z
�z 
ctnt$ o   > ?�y�y 0 
themessage 
theMessage�|  �{   %&% l  C C�x'(�x  '   add recipient   ( �))    a d d   r e c i p i e n t& *+* O   C m,-, I  G l�w�v.
�w .corecrel****      � null�v  . �u/0
�u 
kocl/ m   I L�t
�t 
trcp0 �s12
�s 
insh1 n   O U343  ;   T U4 2  O T�r
�r 
trcp2 �q5�p
�q 
prdt5 K   V h66 �o78
�o 
pnam7 l  Y ^9�n�m9 n   Y ^:;: o   Z ^�l�l 0 thename  ; o   Y Z�k�k 0 therecipient theRecipient�n  �m  8 �j<�i
�j 
radd< l  a f=�h�g= n   a f>?> o   b f�f�f 0 
theaddress 
theAddress? o   a b�e�e 0 therecipient theRecipient�h  �g  �i  �p  - o   C D�d�d 0 
themessage 
theMessage+ @A@ l  n n�cBC�c  B   send the copy   C �DD    s e n d   t h e   c o p yA EFE I  n s�bG�a
�b .emsgsendnull���     bckeG o   n o�`�` 0 
themessage 
theMessage�a  F H�_H l  t t�^�]�\�^  �]  �\  �_   � m    II�                                                                                  emal  alis    F  Macintosh HD               ���4H+     MMail.app                                                        Rn̏�	        ����  	                Applications    ���      ̏�Y       M  #Macintosh HD:Applications: Mail.app     M a i l . a p p    M a c i n t o s h   H D  Applications/Mail.app   / ��   � J�[J l  w w�Z�Y�X�Z  �Y  �X  �[   � KLK l     �W�V�U�W  �V  �U  L MNM l     �T�S�R�T  �S  �R  N OPO l     �QQR�Q  Q + % Purpose:	Open a message given its ID   R �SS J   P u r p o s e : 	 O p e n   a   m e s s a g e   g i v e n   i t s   I DP TUT l     �PVW�P  V   Parameters:   W �XX    P a r a m e t e r s :U YZY l     �O[\�O  [  			theId: message ID   \ �]] ( 	 	 	 t h e I d :   m e s s a g e   I DZ ^_^ l     �N`a�N  `   Returns:	a Message   a �bb &   R e t u r n s : 	 a   M e s s a g e_ cdc l     �M�L�K�M  �L  �K  d efe i    ghg I      �Ji�I�J 0 
getcontent 
getContenti j�Hj o      �G�G 0 theid theID�H  �I  h k     2kk lml l     �F�E�D�F  �E  �D  m non O     0pqp k    /rr sts l   �C�B�A�C  �B  �A  t uvu Q    -wxyw L    zz l   {�@�?{ n    |}| 1    �>
�> 
ctnt} l   ~�=�<~ 6   � n    ��� 4  
 �;�
�; 
mssg� m    �:�: � 1    
�9
�9 
drmb� =    ��� 1    �8
�8 
ID  � o    �7�7 0 theid theID�=  �<  �@  �?  x R      �6��
�6 .ascrerr ****      � ****� o      �5�5 0 	errortext 	errorText� �4��3
�4 
errn� o      �2�2 0 errornumber errorNumber�3  y I  " -�1��0
�1 .sysodlogaskr        TEXT� b   " )��� b   " '��� b   " %��� o   " #�/�/ 0 	errortext 	errorText� m   # $�� ���    [� o   % &�.�. 0 errornumber errorNumber� m   ' (�� ���  ]�0  v ��-� l  . .�,�+�*�,  �+  �*  �-  q m     ���                                                                                  emal  alis    F  Macintosh HD               ���4H+     MMail.app                                                        Rn̏�	        ����  	                Applications    ���      ̏�Y       M  #Macintosh HD:Applications: Mail.app     M a i l . a p p    M a c i n t o s h   H D  Applications/Mail.app   / ��  o ��)� l  1 1�(�'�&�(  �'  �&  �)  f ��� l     �%�$�#�%  �$  �#  � ��� l     �"�!� �"  �!  �   � ��� l     ����  � P J Purpose:	Assemble a list of subjects of each message in the Drafts folder   � ��� �   P u r p o s e : 	 A s s e m b l e   a   l i s t   o f   s u b j e c t s   o f   e a c h   m e s s a g e   i n   t h e   D r a f t s   f o l d e r� ��� l     ����  � X R Returns:	List of Message subjects and IDs; The ID can be used to open the message   � ��� �   R e t u r n s : 	 L i s t   o f   M e s s a g e   s u b j e c t s   a n d   I D s ;   T h e   I D   c a n   b e   u s e d   t o   o p e n   t h e   m e s s a g e� ��� l     ����  �  �  � ��� i    ��� I      ���� 0 gettemplates getTemplates�  �  � k     D�� ��� l     ����  �  �  � ��� O     B��� k    A�� ��� l   ����  �  �  � ��� l   ����  � 7 1 generate a List of drafts in the 'Drafts' folder   � ��� b   g e n e r a t e   a   L i s t   o f   d r a f t s   i n   t h e   ' D r a f t s '   f o l d e r� ��� r    ��� J    ��  � o      �� 0 thesubjectids theSubjectIds� ��� l  	 	����  �  �  � ��� r   	 ��� n  	 ��� 2    �
� 
mssg� 1   	 �

�
 
drmb� o      �	�	 0 themessages theMessages� ��� X    <���� k   ! 7�� ��� l  ! !����  �   Subject [ID]   � ���    S u b j e c t   [ I D ]� ��� r   ! 2��� b   ! 0��� b   ! .��� b   ! (��� l  ! &���� c   ! &��� n   ! $��� 1   " $�
� 
subj� o   ! "�� 0 
themessage 
theMessage� m   $ %�
� 
ctxt�  �  � m   & '�� ���    [� l  ( -��� � c   ( -��� n   ( +��� 1   ) +��
�� 
ID  � o   ( )���� 0 
themessage 
theMessage� m   + ,��
�� 
ctxt�  �   � m   . /�� ���  ]� o      ���� 0 thesubjectid theSubjectId� ���� s   3 7��� o   3 4���� 0 thesubjectid theSubjectId� n      ���  ;   5 6� o   4 5���� 0 thesubjectids theSubjectIds��  � 0 
themessage 
theMessage� o    ���� 0 themessages theMessages� ��� l  = =��������  ��  ��  � ��� L   = ?�� o   = >���� 0 thesubjectids theSubjectIds� ���� l  @ @��������  ��  ��  ��  � m     ���                                                                                  emal  alis    F  Macintosh HD               ���4H+     MMail.app                                                        Rn̏�	        ����  	                Applications    ���      ̏�Y       M  #Macintosh HD:Applications: Mail.app     M a i l . a p p    M a c i n t o s h   H D  Applications/Mail.app   / ��  � ���� l  C C��������  ��  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � > 8 Purpose:	Convert selected Contacts to a List of Records   � ��� p   P u r p o s e : 	 C o n v e r t   s e l e c t e d   C o n t a c t s   t o   a   L i s t   o f   R e c o r d s� ��� l     �� ��      Parameters:    �    P a r a m e t e r s :�  l     ����   _ Y			* categoryName - unused; matches interface of similar function in Outlook Library.scpt    � � 	 	 	 *   c a t e g o r y N a m e   -   u n u s e d ;   m a t c h e s   i n t e r f a c e   o f   s i m i l a r   f u n c t i o n   i n   O u t l o o k   L i b r a r y . s c p t 	 l     ��
��  
 x r Returns:	A Record that contains the distinct list of address types (e.g. 'work') and a List (unsorted) of Records    � �   R e t u r n s : 	 A   R e c o r d   t h a t   c o n t a i n s   t h e   d i s t i n c t   l i s t   o f   a d d r e s s   t y p e s   ( e . g .   ' w o r k ' )   a n d   a   L i s t   ( u n s o r t e d )   o f   R e c o r d s	  l     ����   y s Returns:	A Record that contains the distinct list of address types (e.g. 'work') and a List (unsorted) of Records:    � �   R e t u r n s : 	 A   R e c o r d   t h a t   c o n t a i n s   t h e   d i s t i n c t   l i s t   o f   a d d r e s s   t y p e s   ( e . g .   ' w o r k ' )   a n d   a   L i s t   ( u n s o r t e d )   o f   R e c o r d s :  l     ����    	{    �  	 {  l     ����   ! 		theTypes: {'home','work'}    � 6 	 	 t h e T y p e s :   { ' h o m e ' , ' w o r k ' }  l     ����    		theRecords:{    �    	 	 t h e R e c o r d s : { !"! l     ��#$��  # � � 			{ firstName: Jane, lastName: Doe, displayName: Jane Doe, companyName: Acme, emailAddresses: {{theType:"home", theAddress: jane@doe.name},{theType:"work", theAddress: jane.doe@acme.com}} },   $ �%%�   	 	 	 {   f i r s t N a m e :   J a n e ,   l a s t N a m e :   D o e ,   d i s p l a y N a m e :   J a n e   D o e ,   c o m p a n y N a m e :   A c m e ,   e m a i l A d d r e s s e s :   { { t h e T y p e : " h o m e " ,   t h e A d d r e s s :   j a n e @ d o e . n a m e } , { t h e T y p e : " w o r k " ,   t h e A d d r e s s :   j a n e . d o e @ a c m e . c o m } }   } ," &'& l     ��()��  ( 
 			�   ) �**  	 	 	 &' +,+ l     ��-.��  - � � 			{ firstName: John, lastName: Doe, displayName: John Doe, companyName: Acme, emailAddresses: {{theType:"home", theAddress: john@doe.name},{theType:"work", theAddress: john.doe@acme.com}} }   . �//~   	 	 	 {   f i r s t N a m e :   J o h n ,   l a s t N a m e :   D o e ,   d i s p l a y N a m e :   J o h n   D o e ,   c o m p a n y N a m e :   A c m e ,   e m a i l A d d r e s s e s :   { { t h e T y p e : " h o m e " ,   t h e A d d r e s s :   j o h n @ d o e . n a m e } , { t h e T y p e : " w o r k " ,   t h e A d d r e s s :   j o h n . d o e @ a c m e . c o m } }   }, 010 l     ��23��  2 
 			}   3 �44  	 	 	 }1 565 l     ��78��  7  	}   8 �99  	 }6 :;: l     ��<=��  <   TODOs:	   = �>>    T O D O s : 	; ?@? l     ��AB��  A ' !			* Filter Contacts by groupName   B �CC B 	 	 	 *   F i l t e r   C o n t a c t s   b y   g r o u p N a m e@ DED l     ��FG��  F > 8			* Identify primary email address; not exposed by API?   G �HH p 	 	 	 *   I d e n t i f y   p r i m a r y   e m a i l   a d d r e s s ;   n o t   e x p o s e d   b y   A P I ?E IJI l     ��������  ��  ��  J KLK i    MNM I      ��O���� 0 getcontacts getContactsO P��P o      ���� 0 	groupname 	groupName��  ��  N k     �QQ RSR l     ��������  ��  ��  S TUT O     �VWV k    �XX YZY l   ��������  ��  ��  Z [\[ r    ]^] 4    
��_
�� 
azf5_ l   	`����` c    	aba o    ���� 0 	groupname 	groupNameb m    ��
�� 
ctxt��  ��  ^ o      ���� 0 thegroup theGroup\ cdc r    efe n    ghg 2   ��
�� 
azf4h o    ���� 0 thegroup theGroupf o      ���� 0 thecontacts theContactsd iji l   ��������  ��  ��  j klk l   ��mn��  m 0 *unique list of address types (e.g. 'work')   n �oo T u n i q u e   l i s t   o f   a d d r e s s   t y p e s   ( e . g .   ' w o r k ' )l pqp r    rsr J    ����  s o      ���� 0 	thetypesx 	theTypesXq tut r    vwv J    ����  w o      ���� 0 therecordsx theRecordsXu xyx l   ��������  ��  ��  y z{z X    �|��}| k   - �~~ � l  - -��������  ��  ��  � ��� l  - -������  � : 4 store relevant properties in Record for easy access   � ��� h   s t o r e   r e l e v a n t   p r o p e r t i e s   i n   R e c o r d   f o r   e a s y   a c c e s s� ��� r   - ;��� K   - 9�� ������ 0 	firstname 	firstName� m   . /��
�� 
null� ������ 0 lastname lastName� m   0 1��
�� 
null� ������ 0 displayname displayName� m   2 3��
�� 
null� ������ 0 companyname companyName� m   4 5��
�� 
null� �������  0 emailaddresses emailAddresses� m   6 7��
�� 
null��  � o      ���� 0 	therecord 	theRecord� ��� l  < <��������  ��  ��  � ��� r   < C��� n   < ?��� 1   = ?��
�� 
azf7� o   < =���� 0 
thecontact 
theContact� n      ��� o   @ B���� 0 	firstname 	firstName� o   ? @���� 0 	therecord 	theRecord� ��� r   D K��� n   D G��� 1   E G��
�� 
azf8� o   D E���� 0 
thecontact 
theContact� n      ��� o   H J���� 0 lastname lastName� o   G H���� 0 	therecord 	theRecord� ��� r   L U��� n   L Q��� 1   M Q��
�� 
pnam� o   L M���� 0 
thecontact 
theContact� n      ��� o   R T���� 0 displayname displayName� o   Q R���� 0 	therecord 	theRecord� ��� r   V _��� n   V [��� 1   W [��
�� 
az38� o   V W���� 0 
thecontact 
theContact� n      ��� o   \ ^���� 0 companyname companyName� o   [ \���� 0 	therecord 	theRecord� ��� l  ` `��������  ��  ��  � ��� l  ` `������  � O I get unique list of email addresses; TODO: identify primary email address   � ��� �   g e t   u n i q u e   l i s t   o f   e m a i l   a d d r e s s e s ;   T O D O :   i d e n t i f y   p r i m a r y   e m a i l   a d d r e s s� ��� r   ` d��� J   ` b����  � o      ���� 0 addresslist addressList� ��� r   e i��� J   e g����  � o      ����  0 emailaddresses emailAddresses� ��� l  j j��������  ��  ��  � ��� r   j q��� n   j o��� 2  k o��
�� 
az21� o   j k���� 0 
thecontact 
theContact� o      ���� 0 theaddresses theAddresses� ��� X   r ������ k   � ��� ��� l  � ��������  ��  �  � ��� l  � ��~���~  � / ) build unique type list (for theContacts)   � ��� R   b u i l d   u n i q u e   t y p e   l i s t   ( f o r   t h e C o n t a c t s )� ��� Z   � ����}�|� H   � ��� E   � ���� o   � ��{�{ 0 	thetypesx 	theTypesX� l  � ���z�y� n   � ���� 1   � ��x
�x 
az18� o   � ��w�w 0 
theaddress 
theAddress�z  �y  � r   � ���� b   � ���� o   � ��v�v 0 	thetypesx 	theTypesX� l  � ���u�t� n   � ���� 1   � ��s
�s 
az18� o   � ��r�r 0 
theaddress 
theAddress�u  �t  � o      �q�q 0 	thetypesx 	theTypesX�}  �|  � ��� l  � ��p�o�n�p  �o  �n  � ��� l  � ��m���m  � 0 *build unique address list (for theContact)   � ��� T b u i l d   u n i q u e   a d d r e s s   l i s t   ( f o r   t h e C o n t a c t )� ��� l  � ��l���l  � @ :if addressList does not contain (value of theAddress) then   � ��� t i f   a d d r e s s L i s t   d o e s   n o t   c o n t a i n   ( v a l u e   o f   t h e A d d r e s s )   t h e n� ��� l  � ��k�j�i�k  �j  �i  � ��� l  � ��h���h  � / ) {theType:Home, theAddress:jane@doe.name}   � ��� R   { t h e T y p e : H o m e ,   t h e A d d r e s s : j a n e @ d o e . n a m e }� ��� r   � �� � K   � � �g�g 0 thetype theType l  � ��f�e n   � � 1   � ��d
�d 
az18 o   � ��c�c 0 
theaddress 
theAddress�f  �e   �b�a�b 0 
theaddress 
theAddress l  � ��`�_ n   � �	
	 1   � ��^
�^ 
az17
 o   � ��]�] 0 
theaddress 
theAddress�`  �_  �a    o      �\�\ 0 emailaddress emailAddress�  s   � � o   � ��[�[ 0 emailaddress emailAddress n        ;   � � o   � ��Z�Z  0 emailaddresses emailAddresses  l  � ��Y�X�W�Y  �X  �W    r   � � b   � � o   � ��V�V 0 addresslist addressList l  � ��U�T n   � � 1   � ��S
�S 
az17 o   � ��R�R 0 
theaddress 
theAddress�U  �T   o      �Q�Q 0 addresslist addressList  l  � ��P�P    end if    �    e n d   i f !�O! l  � ��N�M�L�N  �M  �L  �O  �� 0 
theaddress 
theAddress� l  u v"�K�J" o   u v�I�I 0 theaddresses theAddresses�K  �J  � #$# l  � ��H�G�F�H  �G  �F  $ %&% r   � �'(' o   � ��E�E  0 emailaddresses emailAddresses( n      )*) o   � ��D�D  0 emailaddresses emailAddresses* o   � ��C�C 0 	therecord 	theRecord& +,+ s   � �-.- o   � ��B�B 0 	therecord 	theRecord. n      /0/  ;   � �0 o   � ��A�A 0 therecordsx theRecordsX, 1�@1 l  � ��?�>�=�?  �>  �=  �@  �� 0 
thecontact 
theContact} o     !�<�< 0 thecontacts theContacts{ 232 l  � ��;�:�9�;  �:  �9  3 454 l  � ��867�8  6  return theRecords   7 �88 " r e t u r n   t h e R e c o r d s5 9:9 L   � �;; K   � �<< �7=>�7 0 thetypes theTypes= o   � ��6�6 0 	thetypesx 	theTypesX> �5?�4�5 0 
therecords 
theRecords? o   � ��3�3 0 therecordsx theRecordsX�4  : @�2@ l  � ��1�0�/�1  �0  �/  �2  W m     AA�                                                                                  adrb  alis    V  Macintosh HD               ���4H+     MContacts.app                                                     �b�H        ����  	                Applications    ���      �H^`       M  'Macintosh HD:Applications: Contacts.app     C o n t a c t s . a p p    M a c i n t o s h   H D  Applications/Contacts.app   / ��  U B�.B l  � ��-�,�+�-  �,  �+  �.  L CDC l     �*�)�(�*  �)  �(  D EFE l     �'�&�%�'  �&  �%  F GHG l     �$IJ�$  I ' ! Purpose:	Collect names of Groups   J �KK B   P u r p o s e : 	 C o l l e c t   n a m e s   o f   G r o u p sH LML l     �#NO�#  N ) # Returns:	List (unsorted) of Groups   O �PP F   R e t u r n s : 	 L i s t   ( u n s o r t e d )   o f   G r o u p sM QRQ l     �"�!� �"  �!  �   R S�S i    TUT I      ���� 0 	getgroups 	getGroups�  �  U k     0VV WXW l     ����  �  �  X YZY O     .[\[ k    -]] ^_^ l   ����  �  �  _ `a` l   �bc�  b $ generate a List of group names   c �dd < g e n e r a t e   a   L i s t   o f   g r o u p   n a m e sa efe r    ghg J    ��  h o      �� 0 
groupnames 
groupNamesf iji X   	 (k�lk s    #mnm c     opo l   q��q n    rsr 1    �
� 
pnams o    �� 0 	groupitem 	groupItem�  �  p m    �
� 
ctxtn n      tut  ;   ! "u o     !�� 0 
groupnames 
groupNames� 0 	groupitem 	groupIteml 2    �
� 
azf5j vwv l  ) )�
�	��
  �	  �  w xyx L   ) +zz o   ) *�� 0 
groupnames 
groupNamesy {�{ l  , ,����  �  �  �  \ m     ||�                                                                                  adrb  alis    V  Macintosh HD               ���4H+     MContacts.app                                                     �b�H        ����  	                Applications    ���      �H^`       M  'Macintosh HD:Applications: Contacts.app     C o n t a c t s . a p p    M a c i n t o s h   H D  Applications/Contacts.app   / ��  Z }�} l  / /�� ���  �   ��  �  �       ��~�������  ~ �������������� 0 getname getName�� (0 copyandsendmessage copyAndSendMessage�� 0 
getcontent 
getContent�� 0 gettemplates getTemplates�� 0 getcontacts getContacts�� 0 	getgroups 	getGroups �� ����������� 0 getname getName��  ��  �  �  ��� �� �� ����������� (0 copyandsendmessage copyAndSendMessage�� ����� �  �������� 0 theid theID�� 0 
thecontent 
theContent�� 0 therecipient theRecipient��  � �������������� 0 theid theID�� 0 
thecontent 
theContent�� 0 therecipient theRecipient�� 0 advertisement ADVERTISEMENT�� 0 thetemplate theTemplate�� 0 
themessage 
theMessage�  ��� �I����������������������������������������� 0 getname getName
�� 
drmb
�� 
mssg�  
�� 
ID  
�� 
kocl
�� 
bcke
�� 
prdt
�� 
subj
�� 
ctnt
�� 
atts�� �� 
�� .corecrel****      � null
�� 
trcp
�� 
insh
�� 
pnam�� 0 thename  
�� 
radd�� 0 
theaddress 
theAddress
�� .emsgsendnull���     bcke�� y�)j+ %�%E�O� g*�,�k/�[�,\Z�81E�O*�����,����-�� E�O��,�%��,FO� '*�a a *a -6�a �a ,a �a ,�� UO�j OPUOP� ��h���������� 0 
getcontent 
getContent�� ����� �  ���� 0 theid theID��  � �������� 0 theid theID�� 0 	errortext 	errorText�� 0 errornumber errorNumber� �����������������
�� 
drmb
�� 
mssg
�� 
ID  
�� 
ctnt�� 0 	errortext 	errorText� ������
�� 
errn�� 0 errornumber errorNumber��  
�� .sysodlogaskr        TEXT�� 3� - *�,�k/�[�,\Z�81�,EW X  ��%�%�%j 
OPUOP� ������������� 0 gettemplates getTemplates��  ��  � ���������� 0 thesubjectids theSubjectIds�� 0 themessages theMessages�� 0 
themessage 
theMessage�� 0 thesubjectid theSubjectId� �������������������
�� 
drmb
�� 
mssg
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
ID  �� E� ?jvE�O*�,�-E�O *�[��l kh ��,�&�%��,�&%�%E�O��6G[OY��O�OPUOP� ��N���������� 0 getcontacts getContacts�� ����� �  ���� 0 	groupname 	groupName��  � �������������������������� 0 	groupname 	groupName�� 0 thegroup theGroup�� 0 thecontacts theContacts�� 0 	thetypesx 	theTypesX�� 0 therecordsx theRecordsX�� 0 
thecontact 
theContact�� 0 	therecord 	theRecord�� 0 addresslist addressList��  0 emailaddresses emailAddresses�� 0 theaddresses theAddresses�� 0 
theaddress 
theAddress�� 0 emailaddress emailAddress� A��������������������������������������������������
�� 
azf5
�� 
ctxt
�� 
azf4
�� 
kocl
�� 
cobj
�� .corecnte****       ****�� 0 	firstname 	firstName
�� 
null�� 0 lastname lastName�� 0 displayname displayName�� 0 companyname companyName��  0 emailaddresses emailAddresses�� 

�� 
azf7
�� 
azf8
�� 
pnam
�� 
az38
�� 
az21
�� 
az18�� 0 thetype theType�� 0 
theaddress 
theAddress
�� 
az17�� �� 0 thetypes theTypes�� 0 
therecords 
theRecords�� �� �*��&/E�O��-E�OjvE�OjvE�O ��[��l kh �����������E�O��,��,FO��,��,FO�a ,��,FO�a ,��,FOjvE�OjvE�O�a -E�O T�[��l kh 
��a , ��a ,%E�Y hOa �a ,a �a ,a E�O��6GO��a ,%E�OP[OY��O���,FO��6GOP[OY�SOa �a �a OPUOP� ��U���������� 0 	getgroups 	getGroups��  ��  � ����� 0 
groupnames 
groupNames� 0 	groupitem 	groupItem� |�~�}�|�{�z�y
�~ 
azf5
�} 
kocl
�| 
cobj
�{ .corecnte****       ****
�z 
pnam
�y 
ctxt�� 1� +jvE�O *�-[��l kh ��,�&�6G[OY��O�OPUOPascr  ��ޭ