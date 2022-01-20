import {
    ISearchBoxStyleProps,
    ISearchBoxStyles,
    SearchBox
} from "office-ui-fabric-react/lib/SearchBox";
import React from "react";

interface ISearchTextBoxProps {
    placeholderText: string;
    onSearch: (value: string) => void;
    searchBoxStyles: (props: ISearchBoxStyleProps) => Partial<ISearchBoxStyles>;
}

const SearchTextBox = ({
    onSearch,
    placeholderText,
    searchBoxStyles,
}: ISearchTextBoxProps) => {
    return (
        <div className="ms-SearchBox custom-search-box">
            <SearchBox
                styles={searchBoxStyles}
                placeholder={placeholderText}
                onChange={(event: any, value: string) => onSearch(value)}
                maxLength={50}
                title={placeholderText}
            />
        </div>
    );
};

export default SearchTextBox;
