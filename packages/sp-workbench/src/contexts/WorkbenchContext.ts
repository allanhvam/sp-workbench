import { DisplayMode } from "../types/DisplayMode";
import { SPFI, spfi } from "@pnp/sp";
import { createContext } from "react";

type WorkbenchContext = {
    sp: SPFI;
    displayMode: DisplayMode;

    setSP: React.Dispatch<React.SetStateAction<SPFI>>;
    setDisplayMode: React.Dispatch<React.SetStateAction<DisplayMode>>;

    useDefaultSP: ()  => void;
};

export const WorkbenchContext = createContext<WorkbenchContext>({
    sp: spfi().using(),
    displayMode: DisplayMode.Read,

    setSP: () => { },
    setDisplayMode: () => { },

    useDefaultSP: () => { },
});