import * as React from "react";
import { Image, tokens, makeStyles } from "@fluentui/react-components";
import { t } from "i18next";

export interface HeaderProps {
  title: string;
}

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingBottom: "20px",
    paddingTop: "20px",
    backgroundColor: tokens.colorNeutralBackground3,
  },
  message: {
    fontSize: tokens.fontSizeHero900,
    fontWeight: tokens.fontWeightRegular,
    fontColor: tokens.colorNeutralBackgroundStatic,
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const { title } = props; // Removed logo and message
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <h1 className={styles.message}>{t('appTitle')}</h1> {/* Changed to static text "AI Tutor" */}
    </section>
  );
};

export default Header;
