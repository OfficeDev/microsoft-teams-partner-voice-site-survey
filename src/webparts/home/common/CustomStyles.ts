import { FontWeights, getTheme, IIconProps, ITooltipHostStyles, mergeStyleSets } from "@fluentui/react";
import { IButtonStyles } from "@fluentui/react/lib/Button";
import { IDropdownStyles } from "@fluentui/react/lib/Dropdown";
import { IStackProps, IStackStyles } from "@fluentui/react/lib/Stack";
import { ITextFieldStyleProps, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { IPivotStyleProps, IPivotStyles } from "office-ui-fabric-react/lib/components/Pivot";
import { ISearchBoxStyleProps, ISearchBoxStyles } from "office-ui-fabric-react/lib/components/SearchBox";

export const SearchBoxStyles = (props: ISearchBoxStyleProps): Partial<ISearchBoxStyles> => {
    return {
        root: {
            width: 322,
            height: 45,
            fontSize: 18,
            backgroundColor: "#F0F2F7 0% 0% no-repeat padding-box",
            border: "1px solid #7B7CA2",
            opacity: 1,
            selectors: {
                ':after': {
                    border: "2px solid #7B7CA2",
                }
            }
        }
    };
};

export const PivotPanelStyles = (props: IPivotStyleProps): Partial<IPivotStyles> => {
    return {
        root: [
            {
                backgroundColor: 'white',
                borderBottom: 'none',
                height: '60px'
            }
        ],
        link: [
            {
                marginRight: '2px',
                backgroundColor: '#E4E7ED',
                selectors: {
                    ':hover': {
                        backgroundColor: '#E4E7ED',
                    },
                    ':active': {
                        backgroundColor: '#E4E7ED',
                    }
                }
            }
        ],
        linkIsSelected: [
            {
                color: 'black',
                backgroundColor: 'white',
                borderBottom: 'none',
                marginRight: '2px',
                selectors: {
                    ':before': {
                        borderBottom: 'none',
                        backgroundColor: 'white',

                    },
                    ':after': {
                        borderBottom: 'none'
                    },
                    ':hover': {
                        backgroundColor: 'white',
                    },
                    ':active': {
                        backgroundColor: 'white',
                    }


                }
            },
        ]
    };
};

export const StackStyles: Partial<IStackStyles> = { root: { width: "100%", background: "white" } };

export const RibbonStyles: Partial<IStackStyles> = {
    root: {
        width: "100%",
        marginTop: "2%",
    }
};

export const FormProps: Partial<IStackProps> = {
    tokens: { childrenGap: 15 },
    styles: {
        root: {
            width: "100%",
            display: "inline-flex",
            padding: "32px",
        },
    },
};

export const ButtonProps: Partial<IStackProps> = {
    tokens: { childrenGap: 20 },
    styles: {
        root: {
            width: "50%",
            display: "inline-flex",
            paddingTop: "3%"
        },
    },
};

export const StackTokens = { childrenGap: 50 };

export const DropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: "100%" },
};

export const cancelIcon: IIconProps = { iconName: "Cancel" };

export const theme = getTheme();
export const contentStyles = mergeStyleSets({
    container: {
        display: "flex",
        flexFlow: "column nowrap",
        alignItems: "stretch",
    },
    header: [
        // eslint-disable-next-line deprecation/deprecation
        theme.fonts.xLargePlus,
        {
            flex: "1 1 auto",
            backgroundColor: `#243a5e`,
            color: "white",
            display: "flex",
            alignItems: "center",
            fontWeight: FontWeights.semibold,
            padding: "12px 3px 14px 24px",
        },
    ],
    body: {
        flex: "4 4 auto",
        padding: "0px 0px 14px 0px",
        overflowX: "hidden",
        selectors: {
            p: { margin: "14px 0" },
            "p:first-child": { marginTop: 0 },
            "p:last-child": { marginBottom: 0 },
        },
    },
});
export const iconButtonStyles: Partial<IButtonStyles> = {
    root: {
        color: "white",
        marginLeft: "auto",
        marginTop: "4px",
        marginRight: "2px",
    },
    rootHovered: {
        color: theme.palette.neutralDark,
    },
};

export const getNewSiteStyles = (props: ITextFieldStyleProps): Partial<ITextFieldStyles> => {
    return {
        fieldGroup: [
            { width: 300 },
            { display: "inline-flex" },
        ],
        errorMessage: { paddingLeft: "15%" },
        subComponentStyles: {
            label: {
                root: {
                    color: props.theme.palette.themePrimary,
                    display: "inline-flex",
                    marginRight: "10px"
                },
            },
        },
    };
};


export const PopupButtonProps: Partial<IStackProps> = {
    tokens: { childrenGap: 15 },
    styles: {
        root: {
            display: "inline-flex",
            marginTop: "5%",
            marginLeft: "15%"
        },
    },
};

export const backIconClasses = mergeStyleSets({
    icon: {
        fontSize: '22px',
        fontWeight: 'bolder',
        color: '#243a5e',
        cursor: "pointer",
        selectors: {
            ':hover': {
                color: "#9595ef"
            }
        }
    }
});

export const tooltipProps = {
    calloutProps: { gapSpace: 0 },
    styles: {
        content: { color: "black", fontSize: "12px" }
    }
};

export const tooltipStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

export const refreshIconStyles = mergeStyleSets({
    icon: {
        fontSize: '18px',
        fontWeight: 'bolder',
        color: '#1e90ff',
        cursor: 'pointer',
        selectors: {
            ':hover': {
                color: "black"
            }
        }
    }
});

export const cancelBtnStyles: Partial<IButtonStyles> = {
    root: {
        borderColor: "#33344A",
        backgroundColor: "white",
    },
    rootHovered: {
        borderColor: "#33344A",
        backgroundColor: "white",
        color: "#000003"
    },
    rootPressed: {
        borderColor: "#33344A",
        backgroundColor: "white",
        color: "#000003"
    },
    icon: {
        fontSize: "17px",
        fontWeight: "bolder",
        color: "#000003",
        opacity: 1,
        lineHeight: "17px"
    },
    label: {
        font: "normal normal bold 14px/24px Segoe UI",
        letterSpacing: "0px",
        color: "#000003",
        opacity: 1
    }
};

export const newOrgIconclasses = mergeStyleSets({
    Icon: {
        fontSize: '16px',
        color: '#666',
        opacity: 1,
        cursor: 'pointer'
    }
});

export const appHeaderIconClasses = mergeStyleSets({
    Icon: {
        fontSize: '20px',
        color: '#FFFFFF',
        opacity: 1,
        cursor: 'pointer'
    }
});

export const appHeaderCalloutstyles = mergeStyleSets({
    button: {
        width: 130,
    },
    callout: {
        width: 620,
        padding: '10px 24px 20px 24px'
    },
    title: {
        marginBottom: 12,
        fontWeight: FontWeights.bold,
    },
    titlebody: {
        marginBottom: 12,
    },
    titlelink: {
        fontWeight: FontWeights.bold,
        color: "#6264A7"
    },
    link: {
        display: 'block',
        marginBottom: 12,
    },
    linkFont: {
        fontSize: "16px"
    }
});




