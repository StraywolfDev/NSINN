p u b l i c   s t r I n p u t  
 p u b l i c   s t r U s e r  
 p u b l i c   s p e e c h o b j e c t  
 p u b l i c   W s h S h e l l  
 p u b l i c   c h o s e n f a c t  
  
 s t r U s e r   =   C r e a t e O b j e c t ( " W S c r i p t . N e t w o r k " ) . U s e r N a m e  
  
 c a l l   m a i n  
  
 s u b   m a i n  
  
 R a n d o m i z e  
  
 S e t   W s h S h e l l   =   W S c r i p t . C r e a t e O b j e c t ( " W S c r i p t . S h e l l " )  
  
 s e t   o b j F S O = C r e a t e O b j e c t ( " S c r i p t i n g . F i l e S y s t e m O b j e c t " )  
  
 s e t   s p e e c h o b j e c t = c r e a t e o b j e c t ( " s a p i . s p v o i c e " )  
  
 s p e e c h o b j e c t . v o l u m e   =   2 5  
  
 s t r I n p u t   =   U s e r I n p u t (   " G o o g l e   S e a r c h   T e r m : "   )  
  
  
 I f   N O T   I n S t r ( s t r I n p u t ,   " c o n s o l e " )   >   0   T h e n  
  
 I f   N O T   I n S t r ( s t r I n p u t ,   " o p e r a " )   >   0   T h e n  
  
 I f   N O T   I n S t r ( s t r I n p u t ,   " d i s c o r d " )   >   0   T h e n  
  
 I f   I s E m p t y ( s t r I n p u t )   T h e n  
         ' c a n c e l l e d  
         W s h S h e l l . R u n   " t a s k k i l l   / f   / i m   w s c r i p t . e x e "  
 E l s e  
         ' s o m e t h i n g   h a s   e n t e r e d   e v e n   z e r o - l e n g t h  
         S e a r c h O p e r a G X ( s t r I n p u t )  
 E n d   I f  
  
 E l s e  
  
 D i m   o b j S h e l l  
  
 s e t   o b j S h e l l   =   C r e a t e O b j e c t ( " S h e l l . A p p l i c a t i o n " )  
 o b j S h e l l . S h e l l E x e c u t e   " C : \ U s e r s \ "   &   s t r U s e r   &   " \ A p p D a t a \ L o c a l \ D i s c o r d / U p d a t e . e x e "  
 s p e e c h o b j e c t . s p e a k   " O p e n e d   D i s c o r d "  
  
 E n d   I f  
  
 E l s e  
  
 s e t   o b j S h e l l   =   C r e a t e O b j e c t ( " S h e l l . A p p l i c a t i o n " )  
 o b j S h e l l . S h e l l E x e c u t e   " C : \ U s e r s \ "   &   s t r U s e r   &   " \ A p p D a t a \ L o c a l \ P r o g r a m s \ O p e r a   G X \ l a u n c h e r . e x e "  
 s p e e c h o b j e c t . s p e a k   " O p e n e d   O p e r a   G   X "  
  
 E n d   I f  
  
 E l s e  
  
 D i m   o S h e l l  
 S e t   o S h e l l   =   W S c r i p t . C r e a t e O b j e c t   ( " W S C r i p t . s h e l l " )  
 o S h e l l . r u n   " c m d   / K   C D   C : \   &   c o l o r   0 A   &   t i t l e   N S I N N   D e v i c e   C o n s o l e "  
 S e t   o S h e l l   =   N o t h i n g  
 s p e e c h o b j e c t . s p e a k   " O p e n e d   N   S   I   N   N   D e v i c e   C o n   s o l e "  
  
 E n d   I f  
  
 ' M s g B o x   R a n d I t e m F r o m A r r a y ( f a c t s )  
 ' S e a r c h O p e r a G X ( s t r I n p u t )  
  
 c a l l   m a i n  
  
 E n d   S u b  
  
  
 F u n c t i o n   U s e r I n p u t (   m y P r o m p t   )  
  
         I f   U C a s e (   R i g h t (   W S c r i p t . F u l l N a m e ,   1 2   )   )   =   " \ C S C R I P T . E X E "   T h e n  
                  
                 W S c r i p t . S t d O u t . W r i t e   m y P r o m p t   &   "   "  
                 U s e r I n p u t   =   W S c r i p t . S t d I n . R e a d L i n e  
         E l s e  
                  
                 U s e r I n p u t   =   I n p u t B o x (   m y P r o m p t , " N S I N N "   )  
         E n d   I f  
 E n d   F u n c t i o n  
  
 F u n c t i o n   S e a r c h O p e r a G X (   q u e r y   )   ' w i l l   o p e n   n e w   O p e r a G X   w i n d o w   i f   n o t   a l r e a d y   o p e n ,   i f   o p e n   i t   w i l l   o p e n   n e w   t a b   ( b u i l t   i n t o   o p e r a )  
  
 D i m   i U R L    
 D i m   o b j S h e l l  
 D i m   f i x e d q u e r y  
  
 f i x e d q u e r y   =   R e p l a c e ( q u e r y , "   " , " % 2 0 " , 1 , - 1 )  
  
 i U R L   =   " h t t p s : / / w w w . g o o g l e . c o m / s e a r c h ? q = "   &   f i x e d q u e r y  
  
 s e t   o b j S h e l l   =   C r e a t e O b j e c t ( " S h e l l . A p p l i c a t i o n " )  
 o b j S h e l l . S h e l l E x e c u t e   " C : \ U s e r s \ J a c k \ A p p D a t a \ L o c a l \ P r o g r a m s \ O p e r a   G X \ l a u n c h e r . e x e " ,   i U R L ,   " " ,   " " ,   1  
  
 E n d   F u n c t i o n 