import React from "react";
import "./LinearProgressBar.scss";

interface ILinearProgressBarProps {
    completionProgressValue: number;
    maxValue: number;
    status: string;
}

const LinearProgressBar = ({
    completionProgressValue,
    maxValue,
    status
}: ILinearProgressBarProps) => {
    return (
        <React.Fragment>
            <span className="percentage">{completionProgressValue}% </span>
            <progress value={completionProgressValue} max={maxValue}
                className={`${status === 'inprogress' ? 'orangecolor' : status === 'completed' ? 'greencolor' : 'graycolor'}`}
            />
        </React.Fragment>
    );
};

export default LinearProgressBar;
