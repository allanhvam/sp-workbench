import { ReactNode, useCallback, useMemo, useState } from "react";
import { WorkbenchContext } from "../contexts/WorkbenchContext";
import { ISPFXContext, SPFI, SPFx, spfi } from "@pnp/sp";
import { DisplayMode } from "../types/DisplayMode";

type Props = {
  children: ReactNode;

  sp?: SPFI;
};

export function WorkbenchProvider(props: Props) {
  const { children, sp: defaultSP } = props;
  const [displayMode, setDisplayMode] = useState<DisplayMode>(DisplayMode.Read);

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

  const [sp, setSP] = useState<SPFI>(defaultSP ?? spfi().using(SPFx(context)));

  const useDefaultSP = useCallback(() => {
    setSP(defaultSP ?? spfi().using(SPFx(context)));
  }, [defaultSP, setSP]);

  return (
    <WorkbenchContext.Provider
      value={{
        sp,
        displayMode,

        setSP,
        setDisplayMode,

        useDefaultSP,
      }}
    >
      {children}
    </WorkbenchContext.Provider>
  );
}
