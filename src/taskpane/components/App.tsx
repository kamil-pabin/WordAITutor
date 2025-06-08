import React from "react";
import { useState, useEffect } from "react";
import Header from "./Header";
import { makeStyles, shorthands, Button, Dropdown, Option, Spinner, Label, Card, CardHeader, CardPreview, Body1, Body2, Subtitle1, tokens, Subtitle2Stronger, Subtitle2, Slider } from "@fluentui/react-components";
import { SettingsRegular } from "@fluentui/react-icons";
import { useTranslation } from "react-i18next";
import { AnalysisResult } from "../../types";
import "./i18n"; // Initialize i18next

interface AppProps {
  title: string;
}

// Remove StatusMessageState and related logic as i18next handles this

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL, // Replaced shorthands.gap
    ...shorthands.padding(tokens.spacingVerticalL, tokens.spacingHorizontalL),
    backgroundColor: tokens.colorNeutralBackground1,
  },
  card: {
    // Card component has its own padding
  },
  flexColumn: {
    display: "flex",
    flexDirection: "column",
  },
  flexRow: {
    display: "flex",
    flexDirection: "row",
  },
  scoreSection: {
    display: "flex",
    justifyContent: "space-between",
    flexWrap: "wrap",
    ...shorthands.gap(tokens.spacingHorizontalS),
    ...shorthands.padding(tokens.spacingVerticalS, 0),
  },
  scoreItem: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    minWidth: "80px",
    flex: "1",
    textAlign: "left",
    ...shorthands.gap(tokens.spacingVerticalXS),
  },
  statusContainer: {
    display: "flex",
    minWidth: "10px",
    alignItems: "center",
    gap: "8px",
  },
  statusLight: {
    width: "8px",
    height: "8px",
    borderRadius: "50%",
    flexShrink: 0,
  },
  resultText: {
    whiteSpace: "pre-wrap", // To respect newlines from the API response
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1,
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    fontFamily: tokens.fontFamilyBase,
    fontSize: tokens.fontSizeBase300,
  },
  summaryHeader: {
    marginBottom: tokens.spacingVerticalS,
  },
  languageSettingsContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalM,
  },
  dropdownContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingHorizontalXS, // Small gap between label and dropdown
  },
  analysisResultContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
  },
  analysisSection: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    marginTop: tokens.spacingVerticalM,
  },
  cardHeaderText: {
    fontWeight: tokens.fontWeightSemibold,
  }
});

const App: React.FC<AppProps> = () => {
  const { t, i18n } = useTranslation();
  const [statusMessageKey, setStatusMessageKey] = useState<string>("statusIdle");
  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [selectedLanguage, setSelectedLanguage] = useState<string>("en");
  const [detectedLanguage, setDetectedLanguage] = useState<string | null>(null);
  const [statusLight, setStatusLight] = useState<"grey" | "yellow" | "green" | "red">("grey");
  const [rephraseOptions, setRephraseOptions] = useState<{option1: string, option2: string, option3: string} | null>(null);
  const [isRephrasing, setIsRephrasing] = useState(false);
  const [selectedText, setSelectedText] = useState<string>("");
  const [toneValue, setToneValue] = useState<number>(50);
  const [showToneMenu, setShowToneMenu] = useState(false);

  // Close tone menu when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (showToneMenu) {
        const target = event.target as Element;
        if (!target.closest('[data-tone-menu]')) {
          setShowToneMenu(false);
        }
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [showToneMenu]);
  const styles = useStyles();

  const updateStatus = (key: string, light: "grey" | "yellow" | "green" | "red", loading = false) => {
    setStatusMessageKey(key);
    setStatusLight(light);
    setIsLoading(loading);
  };

  React.useEffect(() => {
    const detectLanguageOnLoad = async () => {
      try {
        await Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();
          const documentText = body.text;

          if (documentText && documentText.trim().length > 0) {
            const textSnippet = documentText.substring(0, 500);
            updateStatus("statusAnalyzing", "yellow", true);
            try {
              const response = await fetch('http://localhost:3001/api/detect-language', {
                method: 'POST',
                headers: {
                  'Content-Type': 'application/json',
                },
                body: JSON.stringify({ text: textSnippet }),
              });

              if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || `HTTP error! status: ${response.status}`);
              }

              const data = await response.json();
              if (data.detectedLanguage) {
                const trimmedDetectedLanguage = data.detectedLanguage.trim().toLowerCase();
                setDetectedLanguage(trimmedDetectedLanguage);
                // Map detected language (full name) to language code
                const langMap: { [key: string]: string } = {
                  "english": "en",
                  "spanish": "es",
                  "french": "fr",
                  "german": "de",
                  "japanese": "ja",
                  "polish": "pl"
                };
                const langCode = langMap[trimmedDetectedLanguage] || "en"; // Default to 'en' if not found

                setSelectedLanguage(langCode); // Set selected language to the code
                updateStatus("statusIdle", "green", false);
              } else {
                updateStatus("statusIdle", "yellow", false);
              }
            } catch (fetchError) {
              console.error('Error detecting language:', fetchError);
              updateStatus("statusAnalysisError", "red", false);
            }
          } else {
            updateStatus("statusIdle", "grey", false);
          }
        });
      } catch (error) {
        console.error("Error during language detection Word.run:", error);
        updateStatus("statusAnalysisError", "red", false);
      }
    };

    detectLanguageOnLoad();
  // eslint-disable-next-line react-hooks/exhaustive-deps 
  }, []);

  const handleCheckDocument = async () => {
    console.log("Check Document button clicked");
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        console.log("Document Body Text:");
        console.log(body.text);
        const currentDocumentText = body.text || "<Document is empty or content could not be read>";
        // Use t() to get the full language name for the backend API call
        const fullLanguageName = t(`language${selectedLanguage.toUpperCase()}`);
        updateStatus("statusAnalyzing", "yellow", true);
        setAnalysisResult(null);

        try {
          const response = await fetch('http://localhost:3001/api/analyze', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ text: currentDocumentText, language: fullLanguageName }), // Send full language name
          });

          if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || `HTTP error! status: ${response.status}`);
          }

          const analysis = await response.json();
          console.log('Backend Analysis:', analysis);
          setAnalysisResult(analysis);

          updateStatus("statusAnalysisComplete", "green", false);

        } catch (fetchError) {
          console.error('Error sending data to backend:', fetchError);
          updateStatus("statusAnalysisError", "red", false);
        }
        await context.sync();
      });
    } catch (error) {
      console.error("Error during Word.run:", error);
      updateStatus("statusAnalysisError", "red", false);
      if (error instanceof OfficeExtension.Error && error.debugInfo) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
      }
    }
  };

  const handleRephrase = async () => {
    console.log("Rephrase button clicked");
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        const selectedTextContent = selection.text;
        if (!selectedTextContent || selectedTextContent.trim().length === 0) {
          alert("Please select some text to rephrase.");
          return;
        }

        setSelectedText(selectedTextContent);
        setIsRephrasing(true);
        setRephraseOptions(null);

        const fullLanguageName = t(`language${selectedLanguage.toUpperCase()}`);

        try {
          const response = await fetch('http://localhost:3001/api/rephrase', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ text: selectedTextContent, language: fullLanguageName, tone: toneValue }),
          });

          if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || `HTTP error! status: ${response.status}`);
          }

          const rephraseResult = await response.json();
          console.log('Rephrase Result:', rephraseResult);
          setRephraseOptions(rephraseResult.options);

        } catch (fetchError) {
          console.error('Error sending data to rephrase backend:', fetchError);
          alert('Error rephrasing text. Please try again.');
        } finally {
          setIsRephrasing(false);
        }

        await context.sync();
      });
    } catch (error) {
      console.error("Error during rephrase Word.run:", error);
      setIsRephrasing(false);
      if (error instanceof OfficeExtension.Error && error.debugInfo) {
        console.error("Debug info: " + JSON.stringify(error.debugInfo));
      }
    }
  };

  const replaceSelectedText = async (newText: string) => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(newText, Word.InsertLocation.replace);
        await context.sync();
        setRephraseOptions(null);
        setSelectedText("");
      });
    } catch (error) {
      console.error("Error replacing text:", error);
    }
  };

  const changeUiLanguage = (langCode: string) => {
    i18n.changeLanguage(langCode);
  };

  return (
    <div className={styles.root}>
      <Header title={t("appTitle")} />

      <Card className={styles.card}>
        <CardHeader header={<Subtitle1 className={styles.cardHeaderText}>{t("languageSettingsLabel")}</Subtitle1>} />
        <CardPreview>
          <div className={styles.languageSettingsContainer}>
            <div className={styles.dropdownContainer}>
              <Label htmlFor="uiLanguageDropdown">{t("uiLanguageLabel")}</Label>
              <Dropdown
                style={{maxWidth: "90%"}}
                id="uiLanguageDropdown"
                value={t(`language${i18n.language.toUpperCase()}`)}
                onOptionSelect={(_, data) => changeUiLanguage(data.optionValue as string)}
              >
                <Option value="en">{t("languageEN")}</Option>
                <Option value="es">{t("languageES")}</Option>
                <Option value="pl">{t("languagePL")}</Option>
              </Dropdown>
            </div>
            <div className={styles.dropdownContainer}>
              <Label htmlFor="documentLanguageDropdown">{t("docLanguageLabel")}</Label>
              <Dropdown
                style={{maxWidth: "90%"}}
                id="documentLanguageDropdown"
                value={t(`language${selectedLanguage.toUpperCase()}`)}
                onOptionSelect={(_, data) => {
                  const langCode = data.optionValue as string;
                  setSelectedLanguage(langCode);
                }}
              >
                <Option value="en">{t("languageEN")}</Option>
                <Option value="es">{t("languageES")}</Option>
                <Option value="pl">{t("languagePL")}</Option>
                <Option value="fr">{t("languageFR")}</Option>
                <Option value="de">{t("languageDE")}</Option>
                <Option value="ja">{t("languageJA")}</Option>
              </Dropdown>
            </div>
          </div>
        </CardPreview>
      </Card>

      <div style={{display: "flex", gap: tokens.spacingHorizontalM, alignItems: "center"}}>
        <Button appearance="primary" onClick={handleCheckDocument} disabled={isLoading}>
          {isLoading ? t("checkingInProgress") : t("analyzeButtonText")}
        </Button>
        <div style={{display: "flex", alignItems: "center", gap: tokens.spacingHorizontalXS, position: "relative"}} data-tone-menu>
           <Button appearance="secondary" onClick={handleRephrase} disabled={isRephrasing}>
             {isRephrasing ? t("rephrasingInProgress") : t("rephraseButtonText")}
           </Button>
           <Button 
             appearance="subtle" 
             size="small" 
             icon={<SettingsRegular />}
             onClick={() => setShowToneMenu(!showToneMenu)}
             title={t("toneSettingsTitle")}
             data-tone-menu
           />
           {showToneMenu && (
             <div 
               data-tone-menu
               style={{
                 position: "absolute",
                 top: "100%",
                 right: 0,
                 marginTop: tokens.spacingVerticalXS,
                 padding: tokens.spacingVerticalM,
                 backgroundColor: tokens.colorNeutralBackground1,
                 border: `1px solid ${tokens.colorNeutralStroke1}`,
                 borderRadius: tokens.borderRadiusMedium,
                 boxShadow: tokens.shadow8,
                 zIndex: 1000,
                 minWidth: "200px"
               }}>
              <div style={{marginBottom: tokens.spacingVerticalS, width: '190px'}}>
                <Subtitle2Stronger style={{width: "100%"}}>{t("toneSettingsTitle")}</Subtitle2Stronger>
              </div>
              <div style={{marginBottom: tokens.spacingVerticalS, width: '190px'}}>
                <Body2 style={{width: "100%"}}>{t("toneDescription")}</Body2>
              </div>
              <div style={{display: "flex", marginBottom: tokens.spacingVerticalS, width: '190px'}}>
                <Slider
                  style={{width: "100%"}}
                  value={toneValue}
                  onChange={(_, data) => setToneValue(data.value)}
                  min={0}
                  max={100}
                  step={1}
                />
              </div>
              <div style={{display: "flex", justifyContent: "space-between", fontSize: "12px", color: tokens.colorNeutralForeground2}}>
                <span>{t("toneCasual")}</span>
                <span>{toneValue}</span>
                <span>{t("toneFormal")}</span>
              </div>
            </div>
          )}
        </div>
      </div>

      <Card className={styles.card}>
        <CardHeader 
          header={<Subtitle1 className={styles.cardHeaderText}>{t("statusTitle")}</Subtitle1>} 
          description={
            <div className={styles.statusContainer}>
              <div 
                className={styles.statusLight}
                style={{
                  backgroundColor: statusLight === "green" ? "#28a745" : 
                                 statusLight === "yellow" ? "#ffc107" : 
                                 statusLight === "red" ? "#dc3545" : "#6c757d"
                }}
              />
              {isLoading && <Spinner size="tiny" />}
              <Subtitle2
                style={{
                  color: statusLight === "green" ? tokens.colorPaletteGreenForeground1 : 
                         statusLight === "yellow" ? tokens.colorPaletteYellowForeground1 : 
                         statusLight === "red" ? tokens.colorPaletteRedForeground1 : tokens.colorNeutralForeground2
                }}
              >
                {t(statusMessageKey)}
              </Subtitle2>
            </div>
          }
        />
      </Card>

      {analysisResult && !isLoading && (
        <Card className={styles.card}>
          <CardHeader header={<Subtitle1 className={`${styles.cardHeaderText} ${styles.summaryHeader}`}>{t("analysisSummaryTitle")}</Subtitle1>} />
          <CardPreview style={{padding: tokens.spacingVerticalM}}>
            {analysisResult.analysis.summary && (
              <div className={styles.analysisResultContainer}>
                <Body1 className={styles.resultText}>{analysisResult.analysis.summary}</Body1>
              </div>
            )}
            {analysisResult.analysis.clarityScore && (
                <div className={styles.scoreSection} style={{marginTop: tokens.spacingVerticalM}}>
                    <div className={styles.scoreItem}>
                        <Subtitle2Stronger>{t("clarityLabel")}</Subtitle2Stronger>
                        <Body1>{analysisResult.analysis.clarityScore * 100}/100</Body1>
                    </div>
                    <div className={styles.scoreItem}>
                        <Subtitle2Stronger>{t("sentimentLabel")}</Subtitle2Stronger>
                        <Body1>{t(analysisResult.analysis.sentiment)}</Body1>
                    </div>
                    <div className={styles.scoreItem}>
                        <Subtitle2Stronger>{t("wordCountLabel")}</Subtitle2Stronger>
                        <Body1>{analysisResult.analysis.wordCount}</Body1>
                    </div>
                    <div className={styles.scoreItem}>
                        <Subtitle2Stronger>{t("characterCountLabel")}</Subtitle2Stronger>
                        <Body1>{analysisResult.analysis.characterCount}</Body1>
                    </div>
                </div>
            )}
          </CardPreview>
        </Card>
      )}

      {rephraseOptions && !isRephrasing && (
        <Card className={styles.card}>
          <CardHeader header={<Subtitle1 className={`${styles.cardHeaderText} ${styles.summaryHeader}`}>{t("rephraseOptionsTitle")}</Subtitle1>} />
          <CardPreview style={{padding: tokens.spacingVerticalM}}>
            <div className={styles.analysisResultContainer}>
              {selectedText && (
                  <div style={{display: 'flex', flexDirection: 'column', marginBottom: '10px'}}>
                    <Subtitle2Stronger>{t("originalTextLabel")}:</Subtitle2Stronger>
                    <Body1 className={styles.resultText} style={{fontStyle: "italic", marginTop: tokens.spacingVerticalS, lineHeight: "1.5"}}>{selectedText}</Body1>
                  </div>
                )}
              <div style={{display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM}}>
                <div>
                  <div style={{display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: tokens.spacingVerticalXS}}>
                    <Subtitle2Stronger>{t("option1Label")}:</Subtitle2Stronger>
                     <Button size="small" onClick={() => replaceSelectedText(rephraseOptions.option1)}>{t("useThisButton")}</Button>
                    </div>
                    <Body1 className={styles.resultText}>{rephraseOptions.option1}</Body1>
                  </div>
                  <div>
                    <div style={{display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: tokens.spacingVerticalXS}}>
                      <Subtitle2Stronger>{t("option2Label")}:</Subtitle2Stronger>
                      <Button size="small" onClick={() => replaceSelectedText(rephraseOptions.option2)}>{t("useThisButton")}</Button>
                    </div>
                    <Body1 className={styles.resultText}>{rephraseOptions.option2}</Body1>
                  </div>
                  <div>
                    <div style={{display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: tokens.spacingVerticalXS}}>
                      <Subtitle2Stronger>{t("option3Label")}:</Subtitle2Stronger>
                     <Button size="small" onClick={() => replaceSelectedText(rephraseOptions.option3)}>{t("useThisButton")}</Button>
                  </div>
                  <Body1 className={styles.resultText}>{rephraseOptions.option3}</Body1>
                </div>
              </div>
            </div>
          </CardPreview>
        </Card>
      )}
    </div>
  );
};

export default App;
