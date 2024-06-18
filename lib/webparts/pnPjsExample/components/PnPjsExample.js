var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './PnPjsExample.module.scss';
import { Caching } from "@pnp/queryable";
import { getSP } from "../pnpjsConfig";
import { spfi } from "@pnp/sp";
import { Logger } from "@pnp/logging";
import { Label, PrimaryButton } from '@microsoft/office-ui-fabric-react-bundle';
var PnPjsExample = /** @class */ (function (_super) {
    __extends(PnPjsExample, _super);
    function PnPjsExample(props) {
        var _this = _super.call(this, props) || this;
        _this.LOG_SOURCE = "🅿PnPjsExample";
        _this.LIBRARY_NAME = "Documents";
        _this._readAllFilesSize = function () { return __awaiter(_this, void 0, void 0, function () {
            var spCache, response, items, err_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        spCache = spfi(this._sp).using(Caching({ store: "session" }));
                        return [4 /*yield*/, spCache.web.lists
                                .getByTitle(this.LIBRARY_NAME)
                                .items
                                .select("Id", "Title", "FileLeafRef", "File/Length")
                                .expand("File/Length")()];
                    case 1:
                        response = _a.sent();
                        items = response.map(function (item) {
                            var _a;
                            return {
                                Id: item.Id,
                                Title: item.Title || "Unknown",
                                Size: ((_a = item.File) === null || _a === void 0 ? void 0 : _a.Length) || 0,
                                Name: item.FileLeafRef
                            };
                        });
                        // Add the items to the state
                        this.setState({ items: items });
                        return [3 /*break*/, 3];
                    case 2:
                        err_1 = _a.sent();
                        Logger.write("".concat(this.LOG_SOURCE, " (_readAllFilesSize) - ").concat(JSON.stringify(err_1), " - "), 3 /* LogLevel.Error */);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); };
        _this._updateTitles = function () { return __awaiter(_this, void 0, void 0, function () {
            var _a, batchedSP, execute, items, res_1, i, i, item, err_2;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 6, , 7]);
                        _a = this._sp.batched(), batchedSP = _a[0], execute = _a[1];
                        items = JSON.parse(JSON.stringify(this.state.items));
                        res_1 = [];
                        for (i = 0; i < items.length; i++) {
                            // you need to use .then syntax here as otherwise the application will stop and await the result
                            batchedSP.web.lists
                                .getByTitle(this.LIBRARY_NAME)
                                .items
                                .getById(items[i].Id)
                                .update({ Title: "".concat(items[i].Name, "-Updated") })
                                .then(function (r) { return res_1.push(r); });
                        }
                        // Executes the batched calls
                        return [4 /*yield*/, execute()];
                    case 1:
                        // Executes the batched calls
                        _b.sent();
                        i = 0;
                        _b.label = 2;
                    case 2:
                        if (!(i < res_1.length)) return [3 /*break*/, 5];
                        return [4 /*yield*/, res_1[i].item.select("Id, Title")()];
                    case 3:
                        item = _b.sent();
                        items[i].Name = item.Title;
                        _b.label = 4;
                    case 4:
                        i++;
                        return [3 /*break*/, 2];
                    case 5:
                        //Update the state which rerenders the component
                        this.setState({ items: items });
                        return [3 /*break*/, 7];
                    case 6:
                        err_2 = _b.sent();
                        Logger.write("".concat(this.LOG_SOURCE, " (_updateTitles) - ").concat(JSON.stringify(err_2), " - "), 3 /* LogLevel.Error */);
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        }); };
        // set initial state
        _this.state = {
            items: [],
            errors: []
        };
        _this._sp = getSP();
        return _this;
    }
    PnPjsExample.prototype.componentDidMount = function () {
        // read all file sizes from Documents library
        this._readAllFilesSize();
    };
    PnPjsExample.prototype.render = function () {
        // calculate total of file sizes
        var totalDocs = this.state.items.length > 0
            ? this.state.items.reduce(function (acc, item) {
                return (acc + Number(item.Size));
            }, 0)
            : 0;
        return (React.createElement("div", { className: styles.pnPjsExample },
            React.createElement(Label, { className: styles.title }, "Welcome to PnPJS. This design is inspired by Apple."),
            React.createElement(PrimaryButton, { onClick: this._updateTitles, className: styles.btn }, "Update Item Titles"),
            React.createElement("div", { className: styles.table },
                React.createElement(Label, null, "List of documents:"),
                React.createElement("div", { className: styles.hr }),
                React.createElement("table", { width: "100%" },
                    React.createElement("tr", null,
                        React.createElement("td", null,
                            React.createElement("strong", null, "Title")),
                        React.createElement("td", null,
                            React.createElement("strong", null, "Name")),
                        React.createElement("td", null,
                            React.createElement("strong", null, "Size (KB)"))),
                    this.state.items.map(function (item, idx) {
                        return (React.createElement("tr", { key: idx },
                            React.createElement("td", null, item.Title),
                            React.createElement("td", null, item.Name),
                            React.createElement("td", null, (item.Size / 1024).toFixed(2))));
                    }),
                    React.createElement("tr", null,
                        React.createElement("td", null),
                        React.createElement("td", null,
                            React.createElement("strong", null, "Total:")),
                        React.createElement("td", null,
                            React.createElement("strong", null, (totalDocs / 1024).toFixed(2))))))));
    };
    return PnPjsExample;
}(React.Component));
export default PnPjsExample;
//# sourceMappingURL=PnPjsExample.js.map