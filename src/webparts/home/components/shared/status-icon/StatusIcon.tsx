import React from "react";
import * as Constants from "../../../common/Constants";

interface IStausIconProps {
    iconName: string;
}

const StausIcon = ({ iconName }: IStausIconProps) => {
    const inprogressIcon: string = require("../../../assets/images/inprogress.png");
    const completedIcon: string = require("../../../assets/images//completed.png");
    const notstartedIcon: string = require("../../../assets/images//notstarted.png");

    return (
        <React.Fragment>
            {
                {
                    inprogress: (
                        <img
                            alt={Constants.InProgressText}
                            src={inprogressIcon}
                        />
                    ),
                    completed: (
                        <img
                            alt={Constants.CompletedText}
                            src={completedIcon}
                        />
                    ),
                    notstarted: (
                        <img
                            alt={Constants.NotStartedText}
                            src={notstartedIcon}
                        />
                    ),
                }[iconName]
            }
        </React.Fragment>
    );
};

export default StausIcon;
