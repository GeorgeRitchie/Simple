   M S B A M L   `   `   `  ���� T  PAccountantExcelAutomation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullO,clr-namespace:AccountantExcelAutomation.ViewAccountantExcelAutomation.View  K    EAccountantExcelAutomation.View.WindowsTaskCreateExceptionNotification   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     >local,clr-namespace:AccountantExcelAutomation.View     -   $
��auto��$	��800��  � 
ResizeMode$  NoResize=� � SizeToContent$ Height=� � WindowStartupLocation$ CenterOwner=� � Topmost$
 True�� � WindowStyle$
 None=� � 	Resources    Base �� Property �� Value��%r   &	        �� ���  � ��$ 20����  � ��$ 15(���  � ��$
 Bold"���  � ��$ 0 0 0 0q�.��� -   .+��  J� IsManipulationEnabled$
 True��	 � RowDefinitions	 �� $
��Auto ��� $
��Auto ��� �� �� �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� $��0��$��Center=�$��Center=�$��	10,2,10,2q�$D��>Program cannot create windows task in 'Windows Task Sheduller'������ 
 �� Mode$
 OneWay=� �� Source# [  �� $��1��$��Stretch=�$��Center=�$��	10,2,10,2q�$&�� So you can do that by your self:������ $
 OneWay=�# [  �� -    ��Name$
 path��$��2��$��Stretch=�$��Center=�$��	10,2,10,2q�$��Program path: ��
������ $��0,0,0,0q� �� 
IsReadOnly$
 True�� �� TextWrapping$
 Wrap=�$��20��$
��Auto������ $
 OneWay=�# [  �� -   $ _params��$��3��$��Stretch=�$��Center=�$��	10,2,10,2q�$��Program params: ��
������ $��0,0,0,0q�$
 True��$
 Wrap=�$��20��$
��Auto������ $
 OneWay=�# [  �� -   $ dateTime��$��4��$��Stretch=�$��Center=�$��	10,2,10,2q�$��Program launch time: ��
������ $��0,0,0,0q�$
 True��$
 Wrap=�$��20��$
��Auto������ $
 OneWay=�# [  �� $��5��$��Center=�$��Center=�$��	10,2,10,0q�$��!!!WARNING!!!!��$
��Bold"� �� TextAlignment$ Center=�
���  ������ $
 OneWay=�# [  �� $��6��$��Center=�$��Center=�$��	10,0,10,0q�$W��QIf you make mistake in program path or program params, this program will not run.��$ Center=�
���  ������ $
 OneWay=�# [  �� $��7��$��Center=�$��Center=�$��	10,0,10,2q�$`��ZIf you make mistake in launch time, this program will ignore task and will not send mails.��
���  ��$ Center=����� $
 OneWay=�# [  �� -   $��8��$��Center=�$��Center=�$��10,10,10,12q�$��OK��$��50��$��25������ $
 OneWay=�# [  