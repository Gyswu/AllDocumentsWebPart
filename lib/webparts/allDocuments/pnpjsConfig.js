import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
let _sp;
export const getSP = (context) => {
    if (!_sp && context) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};
//# sourceMappingURL=pnpjsConfig.js.map