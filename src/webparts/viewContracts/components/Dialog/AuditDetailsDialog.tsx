import * as React from "react";
import {
  Dialog,
  DialogFooter,
  DialogType,
} from "office-ui-fabric-react/lib/Dialog";
import { useBoolean } from "@uifabric/react-hooks";
import {
  DefaultButton,
  PrimaryButton,
} from "office-ui-fabric-react/lib/Button";
import styles from "./AuditDetailsDialog.module.scss";
import {
  DetailsList,
  IColumn,
  DetailsListLayoutMode,
  ConstrainMode,
  SelectionMode,
} from "@fluentui/react";

export interface DetailsDialogProps {
  children?: never[];
  audit: any[];
  open: boolean;
  onClose: () => void;
}

export function AuditDetailsDialog(props: DetailsDialogProps) {
  const _col: IColumn[] = [
    {
      key: "Action",
      name: "Action",
      fieldName: "Action",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      data: "string",
      isPadded: true,
    },
    {
      key: "Team",
      name: "Team",
      fieldName: "Team",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      data: "string",
      isPadded: true,
    },
    {
      key: "ActionTakenBy",
      name: "Action taken by",
      fieldName: "ActionTakenBy",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      data: "string",
      isPadded: true,
    },
    {
      key: "ActionTakenOn",
      name: "Action taken on",
      fieldName: "ActionTakenOn",
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      data: "string",
      isPadded: true,
      onRender: (item) =>
        item.ActionTakenOn
    },
  ];


  const { open, onClose, audit } = props;

  const dialogStyles = {
    main: {
      selectors: {
        ["@media (min-width: 480px)"]: {
          width: 1000,
          minWidth: 1000,
          maxWidth: "1000px",
          border: '1px solid #4672c4',
          bordertop: 'none'
        },
      },
    },
  };


  const dialogContentProps = {
    type: DialogType.normal,
    title: "Audit",
  };

  const handleClose = () => () => {
    onClose();
  };

  const modalProps = {
    isBlocking: true,
  };

  const onRenderDetailsHeader = (headerProps, defaultRender) => {
    if (!headerProps || !defaultRender) {
        //technically these may be undefined...
        return null;
    }
    return defaultRender({
        ...headerProps,
        styles: {
            root: {
                selectors: {
                    '.ms-DetailsHeader-cell' : {
                        whiteSpace: 'normal',
                        textOverflow: 'clip',
                        lineHeight: 'normal',
                        background: '#4672c4',
                        borderRight:'1px solid #00008B'                        
                    },
                    ':hover .ms-DetailsHeader-cell' : {
                      whiteSpace: 'normal',
                      textOverflow: 'clip',
                      lineHeight: 'normal',
                      background: '#4672c4',
                      borderRight:'1px solid #00008B'               
                  },
                    '.ms-DetailsHeader-cellTitle': {
                        height: '100%',
                        color:'white',
                        alignItems: 'center',
                        justifyContent: 'center',
                        'font-family': '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                        '-webkit-font-smoothing': 'antialiased',
                        fontSize: '12px',
                        fontWeight: '400'
                    },
                },
            },
        },
    })
}



const onRenderRow = (RowProps, defaultRender) => {
  if (!RowProps || !defaultRender) {
      //technically these may be undefined...
      return null;
  }
  return defaultRender({
      ...RowProps,
      styles: {
          root: {
              selectors: {
                  '.ms-DetailsRow-cell' : {
                      whiteSpace: 'normal',
                      textOverflow: 'clip',
                      lineHeight: 'normal',
                      color:'#4672c4',
                      borderRight:'1px solid #4672c4', 
                      'font-family': '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                      '-webkit-font-smoothing': 'antialiased',
                      fontSize: '12px',
                      fontWeight: '400'                
                  }
              },
          },
      },
  })
}
  

  return (
    <Dialog
     // className={styles.detailsGrid}
      hidden={!open}
      onDismiss={handleClose()}
      dialogContentProps={dialogContentProps}
      styles={dialogStyles}
      modalProps={modalProps}
    >
      <div>
        <DetailsList
          items={audit}
          className={styles.tabContainer}
          columns={_col}
          compact={true}
          layoutMode={DetailsListLayoutMode.justified}
          constrainMode={ConstrainMode.unconstrained}
          onRenderDetailsHeader={onRenderDetailsHeader}
          onRenderRow={onRenderRow}
          isHeaderVisible={true}
          selectionMode={SelectionMode.none}
        ></DetailsList>
      </div>
    </Dialog>
  );
}
