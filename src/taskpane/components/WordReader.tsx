import * as React from "react";
import { Fragment, useState } from "react";
import { Button, Field, makeStyles, tokens } from "@fluentui/react-components";
import axios from "axios";

declare const Office: any;

const useStyles = makeStyles({
  invalidAddressTitle: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
});

const WordReader: React.FC = () => {
  const [invalidEmailAddresses, setInvalidEmailAddresses] = useState<string[]>([]);
  const validateSelection = () => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log("Action failed. Error: " + asyncResult.error.message);
      } else {
        parseData(asyncResult.value);
      }
    });
  };
  function parseData(message) {
    // cleanup previous invalid addresses
    setInvalidEmailAddresses([]);

    // get array of lines
    let lines = message.split("\r");
    // get words
    let words = [];
    for (const line of lines) {
      words.push(...line.split(" "));
    }
    // filter out email addresses
    let emailAddresses = words.filter((word: string) => word.includes("@"));
    console.log("found email addresses", emailAddresses);

    // find out invalid/disposable e-mail addresses
    let promises = [];
    let invalidEmails: string[] = [];
    for (const emailAddress of emailAddresses) {
      console.log("checking", emailAddress);
      promises.push(
        axios
          .get(`https://api.mailcheck.ai/email/${emailAddress}`)
          .then((result: any) => {
            if (result.data["disposable"]) {
              invalidEmails.push(emailAddress);
            }
          })
          .catch(() => {
            invalidEmails.push(emailAddress);
          })
      );
    }
    Promise.all(promises).then(() => {
      setInvalidEmailAddresses([...invalidEmails]);
      console.log("invalid emails", invalidEmails);
    });
  }

  const styles = useStyles();

  return (
    <Fragment>
      <Button
        style={{ marginTop: 1.25 }}
        appearance="primary"
        disabled={false}
        size="large"
        onClick={validateSelection}
      >
        Validate selection
      </Button>
      {typeof invalidEmailAddresses !== "undefined" && invalidEmailAddresses.length > 0 && (
        <Field className={styles.invalidAddressTitle} size="large">
          Invalid e-mail addresses found:
        </Field>
      )}
      {invalidEmailAddresses.map((invalidEmail: string) => (
        <p key={invalidEmail}>{invalidEmail}</p>
      ))}
    </Fragment>
  );
};

export default WordReader;
