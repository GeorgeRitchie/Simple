   M S B A M L   `   `   `  ����T  PAccountantExcelAutomation, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullO,clr-namespace:AccountantExcelAutomation.ViewAccountantExcelAutomation.View  5   	   K    EAccountantExcelAutomation.View.WindowsTaskCreateExceptionNotification   5   	   \ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   >local,clr-namespace:AccountantExcelAutomation.View     5   	   -   5        $
��auto��5      $	��800��6(     � 
ResizeMode$  NoResize=�64    � SizeToContent$ Height=�6J    � WindowStartupLocation$ CenterOwner=�6a    � Topmost$
 True��6�    � WindowStyle$
 None=�6�    � 	Resources 5         Base �� Property �� Value��%�   &	        �� 5	   
   �5
      �� 5       � ��$ 20��6<   6I   �� 5       � ��$ 15(�6:   6G   �� 5       � ��$
 Bold"�6<   6K   �� 5       � ��$ 0 0 0 0q�69   6K   5      5      .��5      � -   5        .+�5   
   �  J� IsManipulationEnabled$
 True��6   	 � RowDefinitions	 5      �� 5      $
��Auto �6    60   �� 5      $
��Auto �6    60   �� 5      6"   �� 5      6"   �� 5      6"   �� 5      $
��Auto �6    60   �� 5      $
��Auto �6    60   �� 5      $
��Auto �6    60   �� 5      $
��Auto �6    60   .+�5!      �� $��0��6   $��Center=�6%   $��Center=�6B   $��	10,2,10,2q�6]   $D��>Program cannot create windows task in 'Windows Task Sheduller'��6p   ��6�   �� 
 �� Mode$
 OneWay=� �� Source# [  �� 5"      $��1��6   $��Stretch=�6%   $��Center=�6C   $��	10,2,10,2q�6^   $&�� So you can do that by your self:��6q   ��6�   �� $
 OneWay=�# [  �� 5#      -   5         ��Name$
 path��5#      $��2��6$   $��Stretch=�61   $��Center=�6O   $��	10,2,10,2q�6j   $��Program path: ��6}   
������ $��0,0,0,0q�6�    �� 
IsReadOnly$
 True��6   �� TextWrapping$
 Wrap=�6  $��20��6(  $
��Auto��67  ��6�   �� $
 OneWay=�# [  67  �� 5$      -   5        $ _params��5$      $��3��6'   $��Stretch=�64   $��Center=�6R   $��	10,2,10,2q�6m   $��Program params: ��6�   
������ $��0,0,0,0q�6�   $
 True��6  $
 Wrap=�6  $��20��6-  $
��Auto��6<  ��6�   �� $
 OneWay=�# [  6<  �� 5%      -   5        $ dateTime��5%      $��4��6(   $��Stretch=�65   $��Center=�6S   $��	10,2,10,2q�6n   $��Program launch time: ��6�   
������ $��0,0,0,0q�6�   $
 True��6  $
 Wrap=�6  $��20��63  $
��Auto��6B  ��6�   �� $
 OneWay=�# [  6B  �� 5&      $��5��6   $��Center=�6%   $��Center=�6B   $��	10,2,10,0q�6]   $��!!!WARNING!!!!��6p   $
��Bold"�6�    �� TextAlignment$ Center=�6�   
���  ����6�   �� $
 OneWay=�# [  �� 5'      $��6��6   $��Center=�6%   $��Center=�6B   $��	10,0,10,0q�6]   $W��QIf you make mistake in program path or program params, this program will not run.��6p   $ Center=�6�   
���  ����6�   �� $
 OneWay=�# [  �� 5(      $��7��6   $��Center=�6%   $��Center=�6B   $��	10,0,10,2q�6]   $`��ZIf you make mistake in launch time, this program will ignore task and will not send mails.��6p   
���  ��$ Center=�6�   ��6�   �� $
 OneWay=�# [  �� 5)      -   5        $��8��5)      $��Center=�6"   $��Center=�6?   $��10,10,10,12q�6Z   $��OK��6o   $��50��6�   $��25��6�   ��6|   �� $
 OneWay=�# [  6�   5*      5+      5,      