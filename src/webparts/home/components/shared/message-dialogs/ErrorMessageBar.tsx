import { MessageBar, MessageBarType } from "@fluentui/react";
import { IconButton } from "@fluentui/react/lib/Button";
import React, { SetStateAction } from "react";

interface Props {
    setError: React.Dispatch<SetStateAction<boolean>>;
    errorMessage: String;
}

const ErrorMessageBar = ({ setError, errorMessage }: Props) => {
    return (
        <React.Fragment>
            <MessageBar
                actions={
                    <div>
                        <IconButton
                            iconProps={{ iconName: "cancel" }}
                            title="Cancel"
                            ariaLabel="Cancel"
                            onClick={() => {
                                setError(false);
                            }}
                        />
                    </div>
                }
                messageBarType={MessageBarType.error}
                isMultiline={false}
                role="status"
            >
                {errorMessage}
            </MessageBar>
        </React.Fragment>
    );
};

export default ErrorMessageBar;
