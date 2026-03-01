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
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
var MyCallendarTeamsWebPart = /** @class */ (function (_super) {
    __extends(MyCallendarTeamsWebPart, _super);
    function MyCallendarTeamsWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._dataSelecionada = new Date();
        _this._eventosAtuais = [];
        return _this;
    }
    MyCallendarTeamsWebPart.prototype.render = function () {
        var corPrincipal = this.properties.corTema || '#0078d4';
        this._renderizarEstruturaBase(corPrincipal);
        this._carregarEventos(corPrincipal).catch(function (err) { return console.error(err); });
    };
    MyCallendarTeamsWebPart.prototype._renderizarEstruturaBase = function (cor) {
        var _this = this;
        var _a, _b;
        var mesOriginal = this._dataSelecionada.toLocaleDateString('pt-BR', { month: 'long' });
        var anoOriginal = this._dataSelecionada.getFullYear();
        var nomeMesFormatado = "".concat(mesOriginal.charAt(0).toUpperCase() + mesOriginal.slice(1), " ").concat(anoOriginal);
        this.domElement.innerHTML = "\n      <div style=\"font-family: 'Segoe UI', system-ui; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);\">\n        <div style=\"display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;\">\n          <div style=\"display: flex; gap: 10px;\">\n            <button id=\"btnAnterior\" title=\"M\u00EAs Anterior\" style=\"cursor:pointer; width:35px; height:35px; background:#f3f2f1; color:".concat(cor, "; border:none; border-radius:50%; font-size:18px; font-weight:bold; display:flex; align-items:center; justify-content:center;\">&#10094;</button>\n            <button id=\"btnProximo\" title=\"Pr\u00F3ximo M\u00EAs\" style=\"cursor:pointer; width:35px; height:35px; background:#f3f2f1; color:").concat(cor, "; border:none; border-radius:50%; font-size:18px; font-weight:bold; display:flex; align-items:center; justify-content:center;\">&#10095;</button>\n          </div>\n          <h3 style=\"margin: 0; font-size: 20px; color: #323130; font-weight: 600;\">").concat(escape(nomeMesFormatado), "</h3>\n          <div style=\"width: 80px;\"></div>\n        </div>\n        <div id=\"calendarioGrid\"></div>\n        <div id=\"painelDetalhes\" style=\"margin-top: 20px; padding: 15px; border-top: 4px solid ").concat(cor, "; background: #faf9f8; display: none; border-radius: 4px;\">\n          <div id=\"listaDetalhes\"></div>\n        </div>\n      </div>");
        (_a = this.domElement.querySelector('#btnAnterior')) === null || _a === void 0 ? void 0 : _a.addEventListener('click', function () { return _this._mudarMes(-1); });
        (_b = this.domElement.querySelector('#btnProximo')) === null || _b === void 0 ? void 0 : _b.addEventListener('click', function () { return _this._mudarMes(1); });
    };
    MyCallendarTeamsWebPart.prototype._mudarMes = function (direcao) {
        this._dataSelecionada.setMonth(this._dataSelecionada.getMonth() + direcao);
        this.render();
    };
    MyCallendarTeamsWebPart.prototype._carregarEventos = function (cor) {
        return __awaiter(this, void 0, void 0, function () {
            var ano, mes, inicio, fim, client, calsResponse, calFeriados, resEventos, todosEventos, resFeriados, feriadosMarcados, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        ano = this._dataSelecionada.getFullYear();
                        mes = this._dataSelecionada.getMonth();
                        inicio = new Date(ano, mes, 1, 0, 0, 0).toISOString();
                        fim = new Date(ano, mes + 1, 0, 23, 59, 59).toISOString();
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 7, , 8]);
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 2:
                        client = _a.sent();
                        return [4 /*yield*/, client.api('/me/calendars').select('id,name').get()];
                    case 3:
                        calsResponse = _a.sent();
                        calFeriados = calsResponse.value.find(function (c) {
                            return c.name.toLowerCase().includes('feriados') || c.name.toLowerCase().includes('holidays');
                        });
                        return [4 /*yield*/, client
                                .api('/me/calendar/calendarView')
                                .header('Prefer', 'outlook.timezone="E. South America Standard Time"')
                                .query({ startDateTime: inicio, endDateTime: fim })
                                .top(999)
                                .get()];
                    case 4:
                        resEventos = _a.sent();
                        todosEventos = resEventos.value || [];
                        if (!calFeriados) return [3 /*break*/, 6];
                        return [4 /*yield*/, client
                                .api("/me/calendars/".concat(calFeriados.id, "/calendarView"))
                                .query({ startDateTime: inicio, endDateTime: fim })
                                .get()];
                    case 5:
                        resFeriados = _a.sent();
                        feriadosMarcados = resFeriados.value.map(function (f) { return (__assign(__assign({}, f), { isFeriado: true })); });
                        todosEventos = __spreadArray(__spreadArray([], todosEventos, true), feriadosMarcados, true);
                        _a.label = 6;
                    case 6:
                        this._eventosAtuais = todosEventos;
                        this._desenharGrade(cor);
                        return [3 /*break*/, 8];
                    case 7:
                        error_1 = _a.sent();
                        console.error("Erro ao carregar dados do Graph:", error_1);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    MyCallendarTeamsWebPart.prototype._desenharGrade = function (cor) {
        var _this = this;
        var ano = this._dataSelecionada.getFullYear();
        var mes = this._dataSelecionada.getMonth();
        var primeiroDiaMes = new Date(ano, mes, 1);
        var ultimoDiaMes = new Date(ano, mes + 1, 0);
        var diasSemana = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
        var html = "<div style=\"display: grid; grid-template-columns: repeat(7, 1fr); width: 100%; background: #edebe9; border-top: 1px solid #edebe9; border-left: 1px solid #edebe9; box-sizing: border-box;\">";
        diasSemana.forEach(function (dia) {
            html += "<div style=\"background: #f3f2f1; padding: 10px 0; text-align: center; font-weight: 600; font-size: 12px; color: #605e5c; border-right: 1px solid #edebe9; border-bottom: 1px solid #edebe9;\">".concat(dia, "</div>");
        });
        for (var i = 0; i < primeiroDiaMes.getDay(); i++) {
            html += '<div style="background: #ffffff; min-height: 100px; border-right: 1px solid #edebe9; border-bottom: 1px solid #edebe9;"></div>';
        }
        var _loop_1 = function (dia) {
            var dataFocoStr = new Date(ano, mes, dia).toLocaleDateString('pt-BR');
            var evs = this_1._eventosAtuais.filter(function (e) { return new Date(e.start.dateTime).toLocaleDateString('pt-BR') === dataFocoStr; });
            var eHoje = new Date().toLocaleDateString('pt-BR') === dataFocoStr;
            html += "\n        <div class=\"dia-calendario\" data-data=\"".concat(dataFocoStr, "\" style=\"background: ").concat(eHoje ? '#f4f4fc' : '#fff', "; min-height: 110px; padding: 5px; border-right: 1px solid #edebe9; border-bottom: 1px solid #edebe9; cursor: pointer; box-sizing: border-box; overflow: hidden;\">\n          <div style=\"font-size: 12px; font-weight: ").concat(eHoje ? '700' : '400', "; color: ").concat(eHoje ? cor : '#323130', ";\">").concat(dia, "</div>\n          <div style=\"margin-top: 5px; display: flex; flex-direction: column; gap: 2px;\">\n            ").concat(evs.slice(0, 3).map(function (e) {
                var hora = new Date(e.start.dateTime).toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
                var corItem = e.isFeriado ? '#8a8886' : cor;
                var textoExibir = e.isFeriado ? escape(e.subject) : "".concat(hora, " ").concat(escape(e.subject));
                return "<div style=\"background: ".concat(corItem, "; color: #fff; font-size: 9px; padding: 1px 3px; border-radius: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\">").concat(textoExibir, "</div>");
            }).join(''), "\n            ").concat(evs.length > 3 ? "<div style=\"font-size: 9px; color: ".concat(cor, "; font-weight: bold;\">+ ").concat(evs.length - 3, " mais</div>") : '', "\n          </div>\n        </div>");
        };
        var this_1 = this;
        for (var dia = 1; dia <= ultimoDiaMes.getDate(); dia++) {
            _loop_1(dia);
        }
        var gridElem = this.domElement.querySelector('#calendarioGrid');
        if (gridElem) {
            gridElem.innerHTML = html + '</div>';
            this.domElement.querySelectorAll('.dia-calendario').forEach(function (el) {
                el.addEventListener('click', function () { return _this._exibirDetalhes(el.getAttribute('data-data') || "", cor); });
            });
        }
    };
    MyCallendarTeamsWebPart.prototype._exibirDetalhes = function (data, cor) {
        var painel = this.domElement.querySelector('#painelDetalhes');
        var lista = this.domElement.querySelector('#listaDetalhes');
        var filtrados = this._eventosAtuais.filter(function (e) { return new Date(e.start.dateTime).toLocaleDateString('pt-BR') === data; });
        if (painel && lista) {
            painel.style.display = 'block';
            lista.innerHTML = "<h4 style=\"margin-top: 0; color: ".concat(cor, ";\">Eventos de ").concat(escape(data), "</h4>");
            if (filtrados.length === 0) {
                lista.innerHTML += "<p style=\"font-size: 14px; color: #605e5c;\">Nenhuma reuni\u00E3o agendada.</p>";
            }
            else {
                filtrados.forEach(function (e) {
                    var inicio = new Date(e.start.dateTime);
                    var fim = new Date(e.end.dateTime);
                    var duracao = Math.round((fim.getTime() - inicio.getTime()) / 60000);
                    var corTexto = e.isFeriado ? '#8a8886' : '#323130';
                    lista.innerHTML += "\n            <div style=\"padding: 10px; border-bottom: 1px solid #edebe9; margin-bottom: 5px; background: #fff; border-radius: 4px;\">\n              <strong style=\"color: ".concat(corTexto, "; font-size: 14px;\">").concat(escape(e.subject), " ").concat(e.isFeriado ? '(Feriado)' : '', "</strong><br/>\n              ").concat(!e.isFeriado ? "\n                <span style=\"font-size: 12px; color: #605e5c;\">\n                  \u23F0 ".concat(inicio.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }), " - ").concat(fim.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }), " \n                  <br/>\u23F3 Dura\u00E7\u00E3o: <strong>").concat(duracao, " minutos</strong>\n                </span>\n              ") : '', "\n            </div>");
                });
            }
            painel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    };
    MyCallendarTeamsWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: { description: "Configure seu Calendário" },
                    groups: [
                        {
                            groupName: "Cores",
                            groupFields: [
                                PropertyPaneTextField('corTema', {
                                    label: 'Código Hex da Cor (Ex: #ff0000)',
                                    description: 'Defina a cor principal da identidade visual'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    Object.defineProperty(MyCallendarTeamsWebPart.prototype, "dataVersion", {
        get: function () { return Version.parse('1.0'); },
        enumerable: false,
        configurable: true
    });
    return MyCallendarTeamsWebPart;
}(BaseClientSideWebPart));
export default MyCallendarTeamsWebPart;
//# sourceMappingURL=MyCallendarTeamsWebPart.js.map