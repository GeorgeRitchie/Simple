   M S B A M L   `   `   `  ���� T  PAccountantExcelAutomation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullO,clr-namespace:AccountantExcelAutomation.ViewAccountantExcelAutomation.View  /    )AccountantExcelAutomation.View.MainWindow   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     >local,clr-namespace:AccountantExcelAutomation.View     X TXceed.Wpf.Toolkit, Version=4.1.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f4<extendentWpf)http://schemas.xceed.com/wpf/xaml/toolkit  -     � Title$  	ExcelAuto��$
��auto��$	��400�� � 
ResizeMode$ NoResize=� � WindowStartupLocation$ CenterScreen=� � SizeToContent$ Height=� � Icon$ /Source/excel.png�� � 	Resources    ToAllElements �� Property �� Value �� Triggers	 P� Property
 P� Value  	ToButtons �� BasedOn  ToSendButtons  ListBoxItems  ToListBoxItemCheckBox �� Binding �� RelativeSource �� Mode �� Path �� Value   ChangeDefaultActionOfButtons �� 
TargetType��%�  &	        &	 _    7   &	 �    7  &	 �    7   &	 S    &	 �    7   �� ���  � ��$ 	10 8 10 8q���  � ��$ Center=���  � ��$ Center=���  � ��$ 30����  �@J�Opacity$	 0.8Q���  � ��$ 15(���  � ��$
 Bold"� P� 	 �@J�	IsFocused$

 True��. ���  �@J�Opacity$ 1Q�P� 	 �@J�IsMouseOver$

 True��. ���  �@J�Opacity$ 1Q���  � ��
 ������� 8   .���  � ��$	 100����  � ��$ 	10 8 10 8q��� 8   .���  � {�$ False���� 8   ���  � ��$ 1 1 1 1q���  � ��$ 0 0 0 0q���  � ��$
 auto����  � ��$ 20����  � ��$ Stretch=���  ��  ��  �� $
 Self=�$ Content� All.a���  �@d�IsThreeState$
 True���� 8   .���  � �� �� " ��.R��� #��z��
������ $��1q�._��� $��Center=�$��Center=� P� 	 �@J�IsMouseOver$

 True��. ���  � ��
 ���� .��� -    �Name$ MainGrid�� J� RenderTransformOrigin$ 0.5,0.5&�����  �� EndPoint$ 0.5,1&� �� 
StartPoint$ 0.5,0&� � SpreadMethod$	 Pad=� � MappingMode$ RelativeToBoundingBox=� �� RelativeTransform V� .~���  �� CenterY$	 0.5Q� �� CenterX$	 0.5Q���  �� CenterY$	 0.5Q� �� CenterX$	 0.5Q���  �� Angle$	 145Q� �� CenterY$	 0.5Q� �� CenterX$	 0.5Q�T�   T� X$  0Q�.��� ! � Color$	! Red��� $! 	#FF0023FF��" � Offset$" 1Q�.+�� -   $ MainOperationsGrid��$y�Visible=�# � RowDefinitions# �� $
��auto ��� $
��auto ��� $
��auto ��� $
��auto �$ � ColumnDefinitions$ �� $��1* ��� $��1* �.+�� $��0��$��0��$��2��$ �� $��1* ��� $
��auto �.+��� -   % ��Name$% Sheets��$��0��$��2��#��[  $��Stretch=�$��
10,20,10,8q�& �� Text$& Select sheet��' � ToolTip$*' $Select sheet with data for receivers��( �� 
IsEditable$
( True��$��2,4,2,2q�$
��True��) �� StaysOpenOnEdit$
) True��* K� ZIndex$* -1��$y�Visible=��� -   + ��Name$+ reloadTheOpenedFile��$��1��#��[ $��
10,20,10,8q�$��30��$��30��$' Reload the opened file��$y�Hidden=����� , �� ImageSource$ , /Source/pngwing.com(1).png��- j� TileMode$
- None=�. j� Stretch$
. Fill=��� -   / ��Name$/ ReceiversList��$��1��$��0��$��2��#��[  $��Stretch=�$	��265��$
��True������ .R�� $ CheckboxGrid��#��[ .+��� -   #��[ $	��340��$��Center=�$��	-5,0,-5,0q����� $
 Name�0 �� Mode$0 TwoWay=�1 d� 	IsChecked1 �� $ 	IsChecked�$0 TwoWay=��� -   $+ SendNow��$��2��$��0��$��Send Now��#��[ �� -	   $+ 	SendLater��$��2��$��1��$��
Send Later��#��[ �� -
   $+ Settings��$��3��$��0��$��2��$��Settings��#��[ $	��295��$
{�True��$��
10,8,10,16q�� -   $ SendLaterGrid��$y�Hidden=�.+�� $��Center=�$��Center=�$ �� $
��auto ��� $
��auto �# �� $
��auto ��� $
��auto �.+�&   Xceed.Wpf.Toolkit.DateTimePicker  -   2  Name$2 DateOfSending��#��[  $��Center=�$
��True��$H' BChoose date and time when mails will be sent to selected receivers��$	��300��$��2��$��Center=�,  &Xceed.Wpf.Toolkit.Primitives.InputBase3   CultureInfo$3  ��4   Text$4 ff��5   TextAlignment$5 Center=��� -   $+ Cancel��$��1��$��0��$��Cancel��#��[ �� -   $+ Confirm��$��1��$��1��$��Confirm��#��[ 