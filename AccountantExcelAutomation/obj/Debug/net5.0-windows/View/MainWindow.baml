   M S B A M L   `   `   `  ����T  PAccountantExcelAutomation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullO,clr-namespace:AccountantExcelAutomation.ViewAccountantExcelAutomation.View  5   	   /    )AccountantExcelAutomation.View.MainWindow   5   	   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   >local,clr-namespace:AccountantExcelAutomation.View     5   	   X TXceed.Wpf.Toolkit, Version=4.1.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f4<extendentWpf)http://schemas.xceed.com/wpf/xaml/toolkit  5   	   -   5          � Title$  	ExcelAuto��5	   	   $
��auto��6   $	��400��6)    � 
ResizeMode$ NoResize=�65    � WindowStartupLocation$ CenterScreen=�6K    � SizeToContent$ Height=�6p    � Icon$ /Source/excel.png��6�    � 	Resources 5
         ToAllElements �� Property �� Value �� Triggers	 P� Property
 P� Value  	ToButtons �� BasedOn  ToSendButtons  ListBoxItems  ToListBoxItemCheckBox �� Binding �� RelativeSource �� Mode �� Path �� Value   ChangeDefaultActionOfButtons �� 
TargetType��%�  &	        &	 j    7   &	 �    7  &	 .    7   &	 5    &	 �    7   �� 5   
   �5      �� 5       � ��$ 	10 8 10 8q�6<   6P   �� 5       � ��$ Center=�6I   6Z   �� 5       � ��$ Center=�6G   6X   �� 5       � ��$ 30��6<   6I   �� 5       �@J�Opacity$	 0.8Q�69   6G   �� 5       � ��$ 15(�6:   6G   �� 5       � ��$
 Bold"�6<   6K    5      P� 5      	 �@J�	IsFocused$

 True��6<   . �5      ��  �@J�Opacity$ 1Q�6=   6I   5      P� 5      	 �@J�IsMouseOver$

 True��6>   . �5      ��  �@J�Opacity$ 1Q�6=   6I   �� 5       � ��
 �����6P   5      5      �� 5    
   8   .�5!      ��  � ��$	 100��67   6E   �� 5"       � ��$ 	10 8 10 8q�68   6L   5#      �� 5%   
   8   .�5&      ��  � {�$ False��67   6G   5'      �� 5)   
   8   �5*      �� 5+       � ��$ 1 1 1 1q�67   6I   �� 5,       � ��$ 0 0 0 0q�68   6J   �� 5-       � ��$
 auto��6;   6J   �� 5.       � ��$ 20��6<   6I   �� 5/       � ��$ Stretch=�6I   6[   51      �� 53   
    54      �� 55       6   ��  �� $
 Self=�$ Content� All6j   .a�56      ��  �@d�IsThreeState$
 True��6>   6M   57      59      �� 5;   
   8   .�5<      ��  � �� 5=      �� 5>      " ��6&   .R�5?      �� #��z��6!   
������ $��1q�6e   ._�5@      �� $��Center=�6/   $��Center=�6L   5A      5B      5D       5E      P� 5F      	 �@J�IsMouseOver$

 True��6>   . �5G      ��  � ��
 ���� 6@   5H      5J      5L      .��5M      � -   5         �Name$ MainGrid��5M       J� RenderTransformOrigin$ 0.5,0.5&�6   ��5N   
   �� 5O       �� EndPoint$ 0.5,1&�6"    �� 
StartPoint$ 0.5,0&�63    � SpreadMethod$	 Pad=�6F    � MappingMode$ RelativeToBoundingBox=�6Y    �� RelativeTransform 5P      V� 5Q      .~�5R      ��  �� CenterY$	 0.5Q�6)    �� CenterX$	 0.5Q�67   �� 5S       �� CenterY$	 0.5Q�6(    �� CenterX$	 0.5Q�66   �� 5T       �� Angle$	 145Q�6*    �� CenterY$	 0.5Q�66    �� CenterX$	 0.5Q�6D   T� 5U        T� X$  0Q�6-   5V      .��5X      � ! � Color$	! Red��6   � 5Y      $! 	#FF0023FF��6   " � Offset$" 1Q�61   5Z      .+�5]   
   � -   5        $ MainOperationsGrid��5]      $y�Visible=�6+   # � RowDefinitions# 5^      �� 5_      $
��auto �6    60   �� 5`      $
��auto �6    60   �� 5a      $
��auto �6    60   �� 5b      $
��auto �6    60   $ � ColumnDefinitions$ 5e      �� 5f      $��1* �6#   60   �� 5g      $��1* �6#   60   .+�5j      � $��0��6   $��0��6    $��2��60   $ 5k      �� 5l      $��1* �6'   64   �� 5m      $
��auto �6'   66   .+�5p      �� -   5        % ��Name$% Sheets��5p      $��0��6+   $��2��6;   #��[  6O   $��Stretch=�6�   $��
10,20,10,8q�6�   & �� Text$& Select sheet��6�   ' � ToolTip$*' $Select sheet with data for receivers��6�   ( �� 
IsEditable$
( True��6�   $��2,4,2,2q�6	  $
��True��6  ) �� StaysOpenOnEdit$
) True��6,  * K� ZIndex$* -1��6n  $y�Visible=�6�  �� 5q      -   5        + ��Name$+ reloadTheOpenedFile��5q      $��1��66   #��[ 6F   $��
10,20,10,8q�6�   $��30��6�   $��30��6�   $' Reload the opened file��6�   $y�Hidden=�6�   ��5r      �� 5s      , �� ImageSource$ , /Source/pngwing.com(1).png��6%   - j� TileMode$
- None=�6N   . j� Stretch$
. Fill=�6^   5u      5v      �� 5x      -   5        / ��Name$/ ReceiversList��5x      $��1��6-   $��0��6:   $��2��6J   #��[  6^   $��Stretch=�6�   $	��265��6�   $
��True��6�   ��5y      �� 5z      .R�5{      � $ CheckboxGrid��6Q   #��[ 6   .+�5|      �� -   5        #��[ 5|   �   $	��340��6�   $��Center=�6�   $��	-5,0,-5,0q�6�   ��6'   �� $
 Name�0 �� Mode$0 TwoWay=�1 d� 	IsChecked1 6R   �� $ 	IsChecked�$0 TwoWay=�6�   5}      5~      5�      �� 5�      -   5        $+ SendNow��5�      $��2��6&   $��0��63   $��Send Now��6C   #��[ 6V   6�   �� 5�      -	   5        $+ 	SendLater��5�      $��2��6(   $��1��65   $��
Send Later��6E   #��[ 6Z   6�   �� 5�      -
   5        $+ Settings��5�      $��3��6'   $��0��64   $��2��6D   $��Settings��6X   #��[ 6k   $	��295��6�   $
{�True��6�   $��
10,8,10,16q�6�   5�      � 5�   
   -   5        $ SendLaterGrid��5�      $y�Hidden=�6&   .+�5�      � $��Center=�6   $��Center=�60   $ 5�      �� 5�      $
��auto �6'   66   �� 5�      $
��auto �6'   66   # 5�      �� 5�      $
��auto �6$   64   �� 5�      $
��auto �6$   64   .+�5�      &   Xceed.Wpf.Toolkit.DateTimePicker  -   5        2  Name$2 DateOfSending��5�   .   #��[  6E   $��Center=�6x   $
��True��6�   $H' BChoose date and time when mails will be sent to selected receivers��6�   $	��300��6�   $��2��6  $��Center=�6  ,  &Xceed.Wpf.Toolkit.Primitives.InputBase3   CultureInfo$3  ��6<  4   Text$4 ff��6K  5   TextAlignment$5 Center=�6U  �� 5�      -   5        $+ Cancel��5�      $��1��6)   $��0��66   $��Cancel��6F   #��[ 6W   6�   �� 5�      -   5        $+ Confirm��5�      $��1��6*   $��1��67   $��Confirm��6G   #��[ 6Y   6�   5�      5�      5�      5�      