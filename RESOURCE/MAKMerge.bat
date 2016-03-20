rc -r -DWIN32_LEAN_AND_MEAN -fo Indx2000.res indx2000_merge.rc
copy indx2000.res ..\APPMAIN
copy indx2000.res ..\APPMAIN_nudi
del indx2000.res

