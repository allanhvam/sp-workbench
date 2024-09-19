import { ToolbarButton } from "@fluentui/react-components";
import clsx from "clsx";
import { ReactNode, useMemo } from "react";
import "@pnp/sp/webs";
import { BookOpenRegular, EditRegular } from "@fluentui/react-icons";
import { useWorkbench } from "../../hooks/useWorkbench";
import { DisplayMode } from "../../types/DisplayMode";
import { useLocation } from "wouter";

type Props = {
  children:
    | ReactNode
    | ((workbench: {
        location: string;
        displayMode: DisplayMode;
      }) => ReactNode);
};

export function SitePage(props: Props) {
  const { children } = props;
  const { displayMode, setDisplayMode } = useWorkbench();
  const [location] = useLocation();

  const editMode = useMemo(() => {
    return displayMode === DisplayMode.Edit;
  }, [displayMode]);

  return (
    <>
      {/* Page Command Bar */}
      <div className="h-10 flex items-center justify-end print:hidden">
        <ToolbarButton
          className={clsx("p-4", {
            "bg-[#037c78] rounded text-white": editMode,
          })}
          icon={editMode ? <BookOpenRegular /> : <EditRegular />}
          onClick={() => {
            setDisplayMode((displayMode: DisplayMode) =>
              displayMode === DisplayMode.Edit
                ? DisplayMode.Read
                : DisplayMode.Edit
            );
          }}
        >
          {editMode ? "Publish" : "Edit"}
        </ToolbarButton>
      </div>

      {/* CanvasComponent */}
      <div>
        {/* Canvas */}
        <div>
          {/* CanvasLayout */}
          <div>
            {/* CanvasZone */}
            <div
              style={{
                padding: "0 16px",
              }}
            >
              {/* CanvasZone-SectionContainer */}
              <div
                style={{
                  display: "flex",
                  maxWidth: "1236px",
                  margin: "auto",
                }}
              >
                {/* CanvasSection */}
                <div
                  style={{
                    minHeight: "auto",
                    width: "100%",
                  }}
                >
                  {/* CanvasControl */}
                  <div
                    style={{
                      margin: "24px 0",
                      padding: "8px",
                    }}
                  >
                    {typeof children === "function"
                      ? children({
                          location,
                          displayMode: editMode
                            ? DisplayMode.Edit
                            : DisplayMode.Read,
                        })
                      : children}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
