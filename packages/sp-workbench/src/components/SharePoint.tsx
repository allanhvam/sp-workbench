import { Title } from "./sp/Title";
import { DisplayMode } from "../types/DisplayMode";
import {
  SearchBox,
  FluentProvider,
  webLightTheme,
} from "@fluentui/react-components";
import { GridDotsFilled, SettingsRegular } from "@fluentui/react-icons";
import { ISPFXContext, SPFx, spfi } from "@pnp/sp";
import { ReactNode, useMemo, useState } from "react";
import { Route, Switch, useLocation } from "wouter";
import { AppBar } from "./sp/AppBar";
import { SitePage } from "./sp/SitePage";
import { ViewLists } from "./sp/ViewLists";
import { WorkbenchContext } from "../contexts/WorkbenchContext";
import { QueryClient } from "@tanstack/react-query";
import { ErrorBoundary } from "./errors/ErrorBoundary";
import { IconButton } from "./sp/IconButton";
import { WorkbenchQueryClientContext } from "../contexts/WorkbenchQueryClientContext";

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
};

export function SharePoint(props: Props) {
  const { children } = props;
  const [editMode, setEditMode] = useState(false);
  const [location, setLocation] = useLocation();

  const context = useMemo<ISPFXContext>(() => {
    return {
      pageContext: {
        web: {
          absoluteUrl: window.location.origin,
        },
        legacyPageContext: {
          formDigestTimeoutSeconds: 60,
          formDigestValue: "digest",
        },
      },
    };
  }, []);

  return (
    <ErrorBoundary resetKeys={["root"]}>
      <FluentProvider theme={webLightTheme}>
        <WorkbenchQueryClientContext.Provider value={queryClient}>
          <WorkbenchContext.Provider
            value={{
              sp: spfi(spfi().using(SPFx(context))),
              displayMode: editMode ? DisplayMode.Edit : DisplayMode.Read,
            }}
          >
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
                  <IconButton
                    className="hover:bg-blue-700"
                    icon={
                      <SettingsRegular className="text-white h-[18px] w-[18px]" />
                    }
                    onClick={() => setLocation("/_layouts/viewlsts.aspx")}
                  />
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

                    <Route path="/_layouts/settings.aspx">
                      Site Settings
                    </Route>

                    <Route path="/_layouts/AdminRecycleBin.aspx">
                      Recycle Bin
                    </Route>

                    <Route>
                      <ErrorBoundary resetKeys={["page"]}>
                        <SitePage editMode={editMode} setEditMode={setEditMode}>
                          {typeof children === "function"
                            ? children({
                              location,
                              displayMode: editMode
                                ? DisplayMode.Edit
                                : DisplayMode.Read,
                            })
                            : children}
                        </SitePage>
                      </ErrorBoundary>
                    </Route>
                  </Switch>
                </div>
              </div>
            </div>
          </WorkbenchContext.Provider>
        </WorkbenchQueryClientContext.Provider>
      </FluentProvider>
    </ErrorBoundary>
  );
}
