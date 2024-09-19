import { useWorkbench } from "../../hooks/useWorkbench";
import { useWorkbenchQuery } from "../../hooks/useWorkbenchQuery";

export function Title() {
  const { sp } = useWorkbench();

  const { data: web } = useWorkbenchQuery(
    {
      queryKey: [sp.web.toUrl(), "webInfo"],
      queryFn: () => sp?.web(),
    },
  );

  return (
    <a
      href={`${web?.ServerRelativeUrl}`}
      className="flex flex-row items-center"
    >
      <div className="mx-4 h-12 w-12 bg-blue-700 text-white font-semibold text-xl flex justify-center items-center">
        {web?.Title?.substring(0, 1)}
      </div>
      <span className="p-4 font-bold text-2xl">{web?.Title}</span>
    </a>
  );
}
