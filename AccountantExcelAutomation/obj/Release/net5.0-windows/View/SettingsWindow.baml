   M S B A M L   `   `   `  ���� T  PAccountantExcelAutomation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullO,clr-namespace:AccountantExcelAutomation.ViewAccountantExcelAutomation.View  3    -AccountantExcelAutomation.View.SettingsWindow   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     >local,clr-namespace:AccountantExcelAutomation.View     -     � Title$  ExcelAuto Settings Window��$	��450��$	��500�� � WindowStartupLocation$ CenterOwner=� � SizeToContent$ Height=� � 
ResizeMode$ NoResize=� � Icon$ /Source/excel.png�� � 	Resources    Base �� Property �� Value  	ToButtons �� BasedOn  ToAllElements	 �� Triggers
 P� Property P� Value  	ToTextBox  ListBoxItems��%q  &	        &	 �     7   &	 �     7   &	 �    7  &	 �    7   �� ���  � ��$ 	10 8 10 8q���  � ��$ Center=���  � ��$ Center=���  � ��$ 20����  � ��$ 15(���  � ��$
 Bold"���  � ��$ 0 0 0 0q��� 8   ���  � ��$	 100���� 8   ���  �@J�Opacity$	 0.8Q�	 P� 
 �@J�	IsFocused$
 True��. ���  �@J�Opacity$ 1Q�P� 
 �@J�IsMouseOver$
 True��. ���  �@J�Opacity$ 1Q���  � ��
 ������� 8   ���  � ��
 �iii���  � ��$ Stretch=���  � ��$ Center=���  � ��$ 5 0 5 0q�	 P� 
 �@J�	IsFocused$
 True��. ���  � ��
 �   �P� 
 �@J�IsMouseOver$
 True��. ���  � ��
 �   ��� 8   ���  � ��$ 5 1 5 1q���  � ��$
 auto����  � ��$ Stretch=���  � ��$ 12(���  � ��$ Normal"�.��� .+�� -    �Name$ MainElementsGrid��$y�Visible=� � ColumnDefinitions �� $��1* ��� $��1* � � RowDefinitions �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �.+��� $��0��$��0��$��2��$=��7Email address of the account in the Outlook application��#��[  � $��1��$��0��$��2��.+��� -    ��Name$ EmailAddress��$��Enter email address...��#��[ � -    �Name$ EmailAddressPopup��$��1��$��2��#��[   � 	StaysOpen$ False��$��Bottom=� � PlacementTarget ��  �� ElementName$ EmailAddress��.���� $��Incorrect email address!��#��� ��$��1,1,1,1q��� $��2��$��0��$��
Mail title��#��[  �� -   $ 	MailTitle��$��2��$��1��$��Enter mail title��#��[ �� $��3��$��0��$��	Mail text��#��[  �� -   $ MailText��$��3��$��1��$��Enter mail text��#��[ �� $��4��$��0��$��Attached file name��#��[  �� -   $ AttachedFileName��$��4��$��1��$��Enter file name��#��[ �� $��5��$��0��$��Attached file extension��#��[  �� -    ��Name$ AttachedFileExtension��$��5��$��1��#��[  �� Text$ Select file extention���� $��6��$��Start trigger value��#��[  �� -	   $ StartTrigger��$��6��$��1��$��Enter start trigger��#��[ �� $��7��$��End trigger value��#��[  �� -
   $ 
EndTrigger��$��7��$��1��$��Enter end trigger��#��[ � -   $ ReceiversListGrid��$��8��$��2�� �� $
��auto ��� $��1* ��� $
��1.5* � �� �� �� .+��� -    ��Name$ Delete��$��0��$��0��$��Delete��#��[ �� -   $ Update��$��1��$��0��$��Update��#��[ �� -   $	 Add��$��2��$��0��$	��Add��#��[ �� -    ��Name$ IDTextBlock��$��0��$��1��$��ID��#��[  $��20��$��0,0,0,0q��� -   $ UpdateTextBlock��$��1��$��1��$
��Name��#��[  $��20��$��15,0,0,0q�$
��Left=��� -   $ AddTextBlock��$��2��$��1��$��Email address��#��[  $��20��$��15,0,0,0q�$
��Left=��� -    ��Name$ ReceiversList��$��2��$��0��$��3��$	��238������ .R�� $	��450�� �� $	��114 ��� $��1* ��� $
��1.5* �.+��� $��0��#��[ $��Center=�����  �� Path$ ID� �� Mode$ TwoWay=��� $��1��#��[ ���� $
 Name�$ TwoWay=��� $��2��#��[ ���� $ EAddress�$ TwoWay=�� -   $ ReceiverEditorGrid��$y�Hidden=�$��Center=� �� $
��auto �$��50���� $
��1.5* ��� $��2* � �� $
��auto ��� $
��auto �.+��� -   $ ReceiverIDEditor��$��0��$��0��$��-�� �� TextAlignment$ Center=����� $ OneWay=� �� Source# [  �� -   $ ReceiverNameEditor��$��1��$��0��$��Enter receiver name...������ $ OneWay=�# [ �� -   $ ReceiverEmailAddress��$��2��$��0��$%��Enter receiver email address...������ $ OneWay=�# [ �� -   $ Operate_button��$��2��$��1��$��Operate��$��0,-2,0,0q����� $ OneWay=�# [ �� -   $ Cancel��$��0��$��2��$��1��$��Cancel��$��0,-2,0,0q����� $ OneWay=�# [ �� -   $ Warnings��$��	0,320,0,0q�$	��Top=�$
��Wrap=�$
��Auto��$ Center=�$y�Hidden=����� $ OneWay=�# [  �� $ �Name or email is not valid���� �� $#�Name must start with a letter���� �� $�Check your email��