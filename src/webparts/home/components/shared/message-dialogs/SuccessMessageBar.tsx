import { MessageBar, MessageBarType } from "@fluentui/react";
import { IconButton } from "@fluentui/react/lib/Button";
import React, { SetStateAction } from "react";

interface Props {
    setSuccess: React.Dispatch<SetStateAction<boolean>>;
    successMessage: String;
}

const SuccessMessageBar = ({ setSuccess, successMessage }: Props) => {
    return (
        <React.Fragment>
            <MessageBar
                actions={
                    <div>
                        <IconButton
                            iconProps={{ iconName: "Cancel" }}
                            title="Cancel"
                            ariaLabel="Cancel"
                            onClick={() => {
                                setSuccess(false);
                            }}
                        />
                    </div>
                }
                messageBarType={MessageBarType.success}
                isMultiline={false}
                role="status"
            >
                {successMessage}
            </MessageBar>
        </React.Fragment>
    );
};

export default SuccessMessageBar;
