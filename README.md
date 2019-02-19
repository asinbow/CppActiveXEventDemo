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
...
Event available: # 3337 RemoteWorkbookNewSheet(Wb,Sh)
Event available: # 3332 RemoteSheetBeforeDelete(Sh)
Event available: # 3333 RemoteSheetPivotTableUpdate(Sh,Target)

@@@@@@@@@@Waiting events, please create a new Excel workbook@@@@@@@@@@
Event triggered: # 1565 NewWorkbook(Wb)
Event triggered: # 1568 WorkbookActivate(Wb)
Event triggered: # 1556 WindowActivate(Wb,Wn)

```

