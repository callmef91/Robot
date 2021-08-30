let o=e=>e,l,c,b,p,h,g,m,y,v,k,u,f,_,x,$;import{css as r,dragger as T,em as C,rem as i,sameBorder as w}from"../utils/style.js";import N from"../../web_modules/classnames.js";import a from"../../web_modules/styled-components.js";const n=a.span.withConfig({displayName:"Table__Dragger",componentId:"he42fd-0"})(["",""],T),E=r(l||(l=o`
    border-spacing: 0;
    ${0};
    border-bottom-left-radius: 0;
    border-bottom-right-radius: 0;

    &.is-droppable-left {
        border-left-color: var(--primary-color);
    }
    &.is-droppable-right {
        border-right-color: var(--primary-color);
    }

    &.sticky {
        overflow: auto;
    }
`),w({radius:!0})),s=r(c||(c=o``)),I=r(b||(b=o``)),d=r(p||(p=o`
    --resizer-pad: ${0};

    margin: 0;
    padding: ${0} ${0};
    border-bottom: 1px solid var(--border-color);
    border-right: 1px solid var(--border-color);
    position: relative;

    &:last-child {
        border-right: none;
    }

    &.is-sticky {
        position: sticky;
    }

    &.is-resizing {
        border-right-color: var(--primary-color);
    }

    &.is-dragging {
        opacity: 0.4;
    }

    &.is-droppable:not(:last-child) {
        border-right-color: var(--primary-color);
    }

    &[data-sticky-td] {
        position: sticky;
    }

    &[data-sticky-last-left-td] {
        box-shadow: 5px 0 3px -3px var(--table-sticky-shadow-color);

        &:not(.is-resizing):not(.is-droppable) {
            border-right-color: var(--table-sticky-shadow-color);
        }
    }
`),i(2),i(15),i(20)),j=r(h||(h=o`
    background-color: var(--table-header-background-color);

    ${0} {
        position: absolute;
        left: var(--resizer-pad);
        top: 50%;
        transform: translateY(-50%);
        color: var(--table-dragger-color);
        opacity: 0;
    }

    &:hover ${0} {
        opacity: 1;
    }
`),n,n),B=r(g||(g=o`
    background-color: var(--table-background-color);

    .tr:hover > & {
        background-color: var(--table-row-hover-background-color);
    }

    .tr:last-child > & {
        border-bottom: none;
    }
`)),D=r(m||(m=o`
    background-color: var(--table-header-background-color);

    .sticky > & {
        top: 0;
        position: sticky;
        z-index: 1;
    }
`)),R=r(y||(y=o`
    .sticky > & {
        position: relative;
        z-index: 0;
    }
`)),t=e=>a.div.attrs(({className:z})=>({className:N([z,e])})),F=t("table")(v||(v=o`
    ${0};
`),E),H=t("thead")(k||(k=o`
    ${0}
    ${0}
`),s,D),S=t("tbody")(u||(u=o`
    ${0}
    ${0}
`),s,R),Y=t("tfoot")(f||(f=o`
    ${0}
`),s),q=t("tr")(_||(_=o`
    ${0}
`),I),A=t("th")(x||(x=o`
    ${0}
    ${0}
`),d,j),G=t("td")($||($=o`
    ${0}
    ${0}
`),d,B),J=a.span.withConfig({displayName:"Table__Resizer",componentId:"he42fd-1"})(["box-sizing:content-box;background-clip:content-box;width:1px;height:100%;position:absolute;top:0;right:calc(var(--resizer-pad) * -1 - 1px);z-index:1;touch-action:none;border-left:var(--resizer-pad) solid transparent;border-right:var(--resizer-pad) solid transparent;&:hover{background-color:var(--primary-color);}&.is-resizing{background-color:var(--primary-color);border-color:var(--primary-color);}.th:last-child > &{display:none;}"]),K=a.span.withConfig({displayName:"Table__Expander",componentId:"he42fd-2"})(["display:inline-block;position:relative;font-size:",";width:1em;height:1em;"," margin-right:0.75em;color:var(--table-expander-color);&::before,&::after{content:'';display:block;background-color:currentColor;position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);}&::before{width:0.5em;height:0.0625em;}&::after{width:0.0625em;height:0.5em;display:",";}&:hover{color:var(--table-expander-hover-color);border-color:var(--table-expander-hover-border-color);}"],C(16),w({color:"var(--table-expander-border-color)",radius:"0.125em"}),e=>e.isExpanded?"none":"block"),L=a.div.withConfig({displayName:"Table__ExpandContainer",componentId:"he42fd-3"})([""," background-color:var(--table-header-background-color);border-right:none;&:last-child{border-bottom:none;}.sticky > .tbody > &{position:sticky;left:0;}"],d);export{F as Table,H as THead,S as TBody,Y as TFoot,q as Tr,A as Th,G as Td,J as Resizer,K as Expander,L as ExpandContainer,n as Dragger};
