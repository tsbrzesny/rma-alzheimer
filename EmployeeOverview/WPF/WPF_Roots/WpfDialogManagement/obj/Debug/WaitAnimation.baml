   M S B A M L   `   `   `  ����$   Technewlogic.WpfDialogManagement4   .Technewlogic.WpfDialogManagement.WaitAnimation   5      W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      $��20��6:   $��20��6E     � Triggers  5      :� 5       :� RoutedEvent$ UserControl.Loaded��6   [�5      �� 5	      .��5
      �� .��5      W�  �� 
TargetName$ rotation��6    �� TargetProperty$ Angle�69   	 W� To$	 360X�5   	    g� Duration$ 0:0:1D�6    g� RepeatBehavior$ Forever��6#   5      5      5      .��5      (� .��5      F� $	��100��6	   $	��100��6    �� Stretch$
 Fill=�6"   
�������$	��0.5��6?   ��5      �� 5       �� 
StartPoint$	 0,0&�6   	 �� EndPoint$		 0,1&�6+   .��5      � 
 � Color$
 	#80000000��6    � Offset$ 0Q�6&   � 5      $
 	#20E0E0E0��6   $ 1Q�6&   5      z�5      �� 5      -   5        rotation�� �� Angle$ 0Q�5   (    �� CenterX$ 50Q�62    �� CenterY$ 50Q�6?   ��5      �� 5       �� GeometryCombineMode$ Exclude=�6    �� 	Geometry1 5       @� 5!       @� RadiusX$ 1Q�6    @� RadiusY$ 1Q�6$    @� Center$ 0.5,0.5&�60    �� 	Geometry2 5#      @� 5$      $
 0.55Q�6   $
 0.55Q�6'   $ 0.5,0.5&�66   5&      5(      5)      5+      