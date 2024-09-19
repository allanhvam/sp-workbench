import { Title } from "./sp/Title";
import { DisplayMode } from "../types/DisplayMode";
import {
  SearchBox,
  FluentProvider,
  webLightTheme,
  Theme,
} from "@fluentui/react-components";
import { GridDotsFilled } from "@fluentui/react-icons";
import { ReactNode } from "react";
import { Route, Switch, useLocation } from "wouter";
import { AppBar } from "./sp/AppBar";
import { SitePage } from "./sp/SitePage";
import { ViewLists } from "./sp/ViewLists";
import { QueryClient } from "@tanstack/react-query";
import { ErrorBoundary } from "./errors/ErrorBoundary";
import { IconButton } from "./sp/IconButton";
import { WorkbenchQueryClientContext } from "../contexts/WorkbenchQueryClientContext";
import { Settings } from "./Settings";
import { WorkbenchProvider } from "../providers/WorkbenchProvider";
import type { SPFI } from "@pnp/sp";

const queryClient = new QueryClient({
  defaultOptions: {
    queries: {
      refetchOnWindowFocus: false,
      useErrorBoundary: true,
      retry: (failureCount, error) => {
        if (failureCount >= 3) {
          return false;
        }
        if (
          error &&
          typeof error === "object" &&
          "status" in error &&
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (error as any).status === 404
        ) {
          // Do not retry on 404
          return false;
        }
        if (error) {
          return true;
        }
        return false;
      },
    },
  },
});

type Props = {
  children:
    | ReactNode
    | ((workbench: {
        location: string;
        displayMode: DisplayMode;
      }) => ReactNode);

  sp?: SPFI;
  theme?: Theme;
};

export function SharePoint(props: Props) {
  const { children, sp, theme } = props;
  const [, setLocation] = useLocation();

  return (
    <ErrorBoundary resetKeys={["root"]}>
      <FluentProvider theme={theme ?? webLightTheme}>
        <WorkbenchQueryClientContext.Provider value={queryClient}>
          <WorkbenchProvider sp={sp}>
            <div className="flex flex-col min-h-screen">
              {/* Office 365 Bar */}
              <div className="h-12 bg-blue-600 flex items-center justify-between print:hidden">
                <IconButton
                  className="hover:bg-blue-700"
                  icon={
                    <GridDotsFilled className="text-white h-[18px] w-[18px]" />
                  }
                  onClick={() => setLocation("/")}
                />
                <SearchBox placeholder={`Search this site`} className="w-96" />
                <div className="px-2">
                  <Settings />
                </div>
              </div>

              <div className="flex flex-row flex-grow">
                <AppBar />
                <div className="w-[calc(100vw-48px)]">
                  {/* SharePoint Header */}
                  <div className="h-20 flex items-center print:hidden">
                    <Title />
                  </div>

                  {/* NOTE: use wouter to void conflicts with react-router-dom */}
                  <Switch>
                    <Route path="/_layouts/viewlsts.aspx">
                      <ViewLists />
                    </Route>

                    <Route path="/_layouts/settings.aspx">Site Settings</Route>

                    <Route path="/_layouts/AdminRecycleBin.aspx">
                      Recycle Bin
                    </Route>

                    <Route>
                      <ErrorBoundary resetKeys={["page"]}>
                        <SitePage>{children}</SitePage>
                      </ErrorBoundary>
                    </Route>
                  </Switch>
                </div>
              </div>
            </div>
          </WorkbenchProvider>
        </WorkbenchQueryClientContext.Provider>
      </FluentProvider>
    </ErrorBoundary>
  );
}
