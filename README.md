CppActiveXEventDemo
---

### Summary

This is a demo project for usage of ActiveX/OLE event in native C++.

Reference: https://github.com/qt/qtactiveqt/blob/v5.12.0/src/activeqt/container/qaxbase.cpp#L2792

### Output

```
Event available: # 1610612736 QueryInterface(riid,ppvObj)
Event available: # 1610612737 AddRef()
Event available: # 1610612738 Release()
Event available: # 1610678272 GetTypeInfoCount(pctinfo)
Event available: # 1610678273 GetTypeInfo(itinfo,lcid,pptinfo)
Event available: # 1610678274 GetIDsOfNames(riid,rgszNames,cNames,lcid,rgdispid)
Event available: # 1610678275 Invoke(dispidMember,riid,lcid,wFlags,pdispparams,pvarResult,pexcepinfo,puArgErr)
Event available: # 1565 NewWorkbook(Wb)
Event available: # 1558 SheetSelectionChange(Sh,Target)
Event available: # 1559 SheetBeforeDoubleClick(Sh,Target,Cancel)
Event available: # 1560 SheetBeforeRightClick(Sh,Target,Cancel)
Event available: # 1561 SheetActivate(Sh)
Event available: # 1562 SheetDeactivate(Sh)
Event available: # 1563 SheetCalculate(Sh)
Event available: # 1564 SheetChange(Sh,Target)
Event available: # 1567 WorkbookOpen(Wb)
Event available: # 1568 WorkbookActivate(Wb)
Event available: # 1569 WorkbookDeactivate(Wb)
Event available: # 1570 WorkbookBeforeClose(Wb,Cancel)
Event available: # 1571 WorkbookBeforeSave(Wb,SaveAsUI,Cancel)
Event available: # 1572 WorkbookBeforePrint(Wb,Cancel)
Event available: # 1573 WorkbookNewSheet(Wb,Sh)
Event available: # 1574 WorkbookAddinInstall(Wb)
Event available: # 1575 WorkbookAddinUninstall(Wb)
Event available: # 1554 WindowResize(Wb,Wn)
Event available: # 1556 WindowActivate(Wb,Wn)
Event available: # 1557 WindowDeactivate(Wb,Wn)
Event available: # 1854 SheetFollowHyperlink(Sh,Target)
Event available: # 2157 SheetPivotTableUpdate(Sh,Target)
Event available: # 2160 WorkbookPivotTableCloseConnection(Wb,Target)
Event available: # 2161 WorkbookPivotTableOpenConnection(Wb,Target)
Event available: # 2289 WorkbookSync(Wb,SyncEventType)
Event available: # 2290 WorkbookBeforeXmlImport(Wb,Map,Url,IsRefresh,Cancel)
Event available: # 2291 WorkbookAfterXmlImport(Wb,Map,IsRefresh,Result)
Event available: # 2292 WorkbookBeforeXmlExport(Wb,Map,Url,Cancel)
Event available: # 2293 WorkbookAfterXmlExport(Wb,Map,Url,Result)
Event available: # 2611 WorkbookRowsetComplete(Wb,Description,Sheet,Success)
Event available: # 2612 AfterCalculate()
Event available: # 2895 SheetPivotTableAfterValueChange(Sh,TargetPivotTable,TargetRange)
Event available: # 2896 SheetPivotTableBeforeAllocateChanges(Sh,TargetPivotTable,ValueChangeStart,ValueChangeEnd,Cancel)
Event available: # 2897 SheetPivotTableBeforeCommitChanges(Sh,TargetPivotTable,ValueChangeStart,ValueChangeEnd,Cancel)
Event available: # 2898 SheetPivotTableBeforeDiscardChanges(Sh,TargetPivotTable,ValueChangeStart,ValueChangeEnd)
Event available: # 2903 ProtectedViewWindowOpen(Pvw)
Event available: # 2905 ProtectedViewWindowBeforeEdit(Pvw,Cancel)
Event available: # 2906 ProtectedViewWindowBeforeClose(Pvw,Reason,Cancel)
Event available: # 2908 ProtectedViewWindowResize(Pvw)
Event available: # 2909 ProtectedViewWindowActivate(Pvw)
Event available: # 2910 ProtectedViewWindowDeactivate(Pvw)
Event available: # 2911 WorkbookAfterSave(Wb,Success)
Event available: # 2912 WorkbookNewChart(Wb,Ch)
Event available: # 3075 SheetLensGalleryRenderComplete(Sh)
Event available: # 3076 SheetTableUpdate(Sh,Target)
Event available: # 3080 WorkbookModelChange(Wb,Changes)
Event available: # 3079 SheetBeforeDelete(Sh)
Event available: # 3335 WorkbookBeforeRemoteChange(Wb)
Event available: # 3336 WorkbookAfterRemoteChange(Wb)
Event available: # 3329 RemoteSheetChange(Sh,Target)
Event available: # 3337 RemoteWorkbookNewSheet(Wb,Sh)
Event available: # 3338 RemoteWorkbookNewChart(Wb,Ch)
Event available: # 3332 RemoteSheetBeforeDelete(Sh)
Event available: # 3333 RemoteSheetPivotTableUpdate(Sh,Target)

@@@@@@@@@@Waiting events, please create a new Excel workbook@@@@@@@@@@
Event triggered: # 1565 NewWorkbook(Wb)
Event triggered: # 1568 WorkbookActivate(Wb)
Event triggered: # 1556 WindowActivate(Wb,Wn)

```

