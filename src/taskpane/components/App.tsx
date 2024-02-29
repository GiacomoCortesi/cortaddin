import * as React from "react";
import Header from "./Header";
import AddinDescription, { AppDescriptionItem } from "./AddinDescription";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular } from "@fluentui/react-icons";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props: AppProps) => {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: AppDescriptionItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Validate e-mail addresses directly in MS Office",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Leverage mailcheck public REST API for validation",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/dp.png" title={props.title} message="e-mail validator Addin" />
      <AddinDescription
        message="This addin allows to validate and add e-mail addresses into a word document"
        items={listItems}
      />
      <TextInsertion />
    </div>
  );
};

export default App;
