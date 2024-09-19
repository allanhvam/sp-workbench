import {
  Button,
  Drawer,
  DrawerHeader,
  DrawerHeaderTitle,
  DrawerBody,
  Link,
  Text,
} from "@fluentui/react-components";
import { IconButton } from "./sp/IconButton";
import { SettingsRegular, Dismiss24Regular } from "@fluentui/react-icons";
import { useLocation } from "wouter";
import { useState } from "react";
import { useWorkbench } from "../hooks/useWorkbench";
import { SPBrowser, spfi } from "@pnp/sp";
import { MSAL, MSALOptions } from "../msal";

export function Settings() {
  const [, setLocation] = useLocation();
  const [isOpen, setIsOpen] = useState(false);
  const workbench = useWorkbench();

  return (
    <>
      <Drawer
        type={"overlay"}
        separator
        open={isOpen}
        position="end"
        onOpenChange={(_, { open }) => setIsOpen(open)}
      >
        <DrawerHeader>
          <DrawerHeaderTitle
            action={
              <Button
                appearance="subtle"
                aria-label="Close"
                icon={<Dismiss24Regular />}
                onClick={() => setIsOpen(false)}
              />
            }
          >
            Settings
          </DrawerHeaderTitle>
        </DrawerHeader>

        <DrawerBody>
          <>
            <Link
              href="/_layouts/viewlsts.aspx"
              onClick={(e) => {
                e.preventDefault();
                setLocation("/_layouts/viewlsts.aspx");
                setIsOpen(false);
              }}
            >
              Site contents
            </Link>

            <div className="pt-8">
              <Text size={500} weight="semibold">
                Workbench
              </Text>

              <div className="flex flex-col">
                <Link
                  onClick={(e) => {
                    e.preventDefault();

                    workbench.useDefaultSP();

                    setLocation("/");
                    setIsOpen(false);
                  }}
                >
                  Use Default
                </Link>
                <Link
                  onClick={(e) => {
                    e.preventDefault();

                    // TODO: Dialog with guide and form.

                    const site = prompt("Site", "");
                    const tenantId = prompt("Tenant Id", "");
                    const clientId = prompt("Client Id", "");

                    if (!site || !tenantId || !clientId) {
                      return;
                    }

                    const { hostname, pathname } = new URL(site);
                    const tenant = hostname.split(".")[0];

                    const { origin } = new URL(window.location.href);

                    const configuration: MSALOptions = {
                      configuration: {
                        auth: {
                          clientId,
                          authority: `https://login.microsoftonline.com/${tenantId}`,
                          redirectUri: origin,
                        },
                        cache: {
                          cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
                          storeAuthStateInCookie: true, // Recommended to avoid certain IE/Edge issues.
                        },
                      },
                      authParams: {
                        forceRefresh: false,
                        scopes: [`https://${tenant}.sharepoint.com/.default`],
                      },
                    };

                    workbench.setSP(
                      spfi(site).using(SPBrowser(), MSAL(configuration))
                    );

                    setLocation(pathname);
                    setIsOpen(false);
                  }}
                >
                  Use Site
                </Link>
              </div>
            </div>
          </>
        </DrawerBody>
      </Drawer>

      <IconButton
        className="hover:bg-blue-700"
        icon={<SettingsRegular className="text-white h-[18px] w-[18px]" />}
        onClick={() => setIsOpen(!isOpen)}
      />
    </>
  );
}
