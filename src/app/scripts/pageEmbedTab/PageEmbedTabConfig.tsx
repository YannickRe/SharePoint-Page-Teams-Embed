import * as React from "react";
import {
    Input,
    getContext,
    TeamsThemeContext
} from "msteams-ui-components-react";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

export interface IPageEmbedTabConfigState extends ITeamsBaseComponentState {
    entityId: string;
    tabName: string;
    relativeUrl: string;
    teamSiteDomain: string;
}

export interface IPageEmbedTabConfigProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of PageEmbed configuration page
 */
export class PageEmbedTabConfig  extends TeamsBaseComponent<IPageEmbedTabConfigProps, IPageEmbedTabConfigState> {

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context: microsoftTeams.Context) => {
                this.updateTheme(context.theme);

                const teamSiteDomain = (context as any).teamSiteDomain;

                this.setState({
                    entityId: context.entityId,
                    teamSiteDomain,
                    tabName: "SharePoint Page"
                });

                microsoftTeams.settings.getSettings((instanceSettings: microsoftTeams.settings.Settings) => {
                    this.setState({
                        relativeUrl: instanceSettings.contentUrl.replace(`https://${teamSiteDomain}/_layouts/15/teamslogon.aspx?spfx=true&dest=`, "")
                    });
                });
            });

            microsoftTeams.settings.registerOnSaveHandler((saveEvent: microsoftTeams.settings.SaveEvent) => {
                const url = `https://${this.state.teamSiteDomain}/_layouts/15/teamslogon.aspx?spfx=true&dest=${this.state.relativeUrl}`;
                microsoftTeams.settings.setSettings({
                    contentUrl: url,
                    websiteUrl: url,
                    suggestedDisplayName: this.state.tabName,
                    entityId: "sharepointPageEmbed"
                });
                saveEvent.notifySuccess();
            });
        }
    }

    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes } = font;
        const styles = {
            section: { ...sizes.base, marginTop: rem(3.2) },
        };

        if (!this.state.teamSiteDomain) {
           return null;
        }

        this.validateSettings();

        return (
            <TeamsThemeContext.Provider value={context}>
                {!this.state.entityId && (<section style={styles.section}>
                    <Input
                        autoFocus
                        label="Tab name"
                        errorLabel={!this.state.tabName ? "Name is required" : undefined}
                        value={this.state.tabName}
                        onChange={(e) => {
                            this.setState({
                                tabName: e.target.value
                            });
                        }}
                        required />
                    </section>)}
                <section style={styles.section}>
                    <Input
                        placeholder="/sites/TheIntranet/SitePages/Home.aspx"
                        label="Relative URL to a SharePoint page"
                        errorLabel={!this.state.relativeUrl ? "This value is required" : undefined}
                        value={this.state.relativeUrl}
                        onChange={(e) => {
                            this.setState({
                                relativeUrl: e.target.value
                            });
                        }}
                        required />
                </section>
            </TeamsThemeContext.Provider>
        );
    }

    private validateSettings(): void {
        this.setValidityState(false);
        if (this.state.tabName && this.state.relativeUrl) {
            this.setValidityState(true);
        }
    }
}
