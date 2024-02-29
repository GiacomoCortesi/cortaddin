import * as React from "react";
import { DismissRegular } from "@fluentui/react-icons";
import {
  MessageBar,
  MessageBarActions,
  MessageBarTitle,
  MessageBarBody,
  Button,
  MessageBarGroup,
} from "@fluentui/react-components";
interface props {
  open: boolean;
  handleClose: any;
}

export const EmailValidationMessage = (props: props) => {
  const { open, handleClose } = props;

  return (
    <MessageBarGroup>
      <MessageBar style={{ display: !open && "none" }}>
        <MessageBarBody>
          <MessageBarTitle>Error</MessageBarTitle>
          Provided e-mail is invalid
        </MessageBarBody>
        <MessageBarActions
          containerAction={
            <Button
              onClick={() => handleClose()}
              aria-label="dismiss"
              appearance="transparent"
              icon={<DismissRegular />}
            />
          }
        ></MessageBarActions>
      </MessageBar>
    </MessageBarGroup>
  );
};
