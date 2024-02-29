import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import insertText from "../office-document";
import axios from "axios";
import { EmailValidationMessage } from "./EmailValidationMessage";
import WordReader from "./WordReader";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const TextInsertion: React.FC = () => {
  const [text, setText] = useState<string>("");
  const [showMessage, setShowMessage] = useState<boolean>(false);

  const handleMessageClose = () => {
    setShowMessage(false);
  };
  const handleTextInsertion = async () => {
    await insertText(text);
  };

  const handleTextValidation = () => {
    axios
      .get(`https://api.mailcheck.ai/email/${text}`)
      .then((result: any) => {
        console.log(result);
        if (result.data["disposable"]) {
          console.log("disposable e-mail");
          setShowMessage(true);
        } else {
          handleTextInsertion();
        }
      })
      .catch(() => {
        console.log("invalid address");
        setShowMessage(true);
      });
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Insert valid email address.">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click the button to insert the email address.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextValidation}>
        Insert e-mail address
      </Button>
      <EmailValidationMessage open={showMessage} handleClose={handleMessageClose} />
      <WordReader />
    </div>
  );
};

export default TextInsertion;
