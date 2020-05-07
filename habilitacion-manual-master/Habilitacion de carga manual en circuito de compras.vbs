/**** actualiza la parameters******/

 update CWPARAMETERS set
 NORMALVALUE = 'S'
 WHERE
 paramname = 'CO_CARGAMANUALITM'

 update CWPARAMETERS set
 NORMALVALUE = 'S'
 WHERE
 paramname = 'FC_CARGAMANUALITM'

 /***** Actualiza circuito que quiero implementar con carga manual *****/
 update FCTCIH set
 FCTCIH_CRGMAN = 'S'
 where
 FCTCIH_circom = '0300' and FCTCIH_cirapl = '0250'
 
 update cotcih set
 COTCIH_CRGMAN = 'S'
 where
 cotcih_circom = '0305' and cotcih_cirapl = '0200'

/**** pone visibles campos APL *******/

UPDATE CWTMFIELDS SET VirtualMode = ''
WHERE TABLENAME = 'CORMVI'
AND FIELDNAME IN ('CORMVI_MODAPL', 'CORMVI_CODAPL', 'CORMVI_NROAPL', 'CORMVI_ITMAPL')

UPDATE CWTMFIELDS SET VirtualMode = ''
WHERE TABLENAME = 'FCRMVI'
AND FIELDNAME IN ('FCRMVI_MODAPL', 'FCRMVI_CODAPL', 'FCRMVI_NROAPL', 'FCRMVI_ITMAPL')