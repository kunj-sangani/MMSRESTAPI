import * as React from 'react';
import { Panel, IPanelProps } from 'office-ui-fabric-react/lib/Panel';
import { IAllGroups } from "./TermstoreInterfaces";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton } from 'office-ui-fabric-react/lib/components/Button/PrimaryButton/PrimaryButton';

export interface IGroupPanelProps {
    isOpenGroup: boolean;
    dismissPanelGroup: any;
    selectedpanelGroup?: IAllGroups;
    onUpdate: any;
    onAdd: any;
}

export interface IGroupPanelState {
    groupName: string;
    groupDescription: string;
}

export default class GroupPanel extends React.Component<IGroupPanelProps, IGroupPanelState> {

    public componentDidUpdate(prevProps) {
        if (this.props.selectedpanelGroup !== prevProps.selectedpanelGroup) {
            this.setState({
                groupName: this.props.selectedpanelGroup.name,
                groupDescription: this.props.selectedpanelGroup.description
            });
        }
    }

    constructor(props: IGroupPanelProps, state: IGroupPanelState) {
        super(props);
        this.state = {
            groupName: this.props.selectedpanelGroup ? this.props.selectedpanelGroup.name : "",
            groupDescription: this.props.selectedpanelGroup ? this.props.selectedpanelGroup.description : "",
        };
    }

    private onChangegroupName = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        this.setState({
            groupName: newValue
        });
    }

    private onChangegroupDescription = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        this.setState({
            groupDescription: newValue
        });
    }

    public render(): React.ReactElement<IGroupPanelProps> {
        return (
            <Panel
                isOpen={this.props.isOpenGroup}
                onDismiss={this.props.dismissPanelGroup}
                headerText={this.props.selectedpanelGroup ? "Edit Group" : "New Group"}
                closeButtonAriaLabel="Close"
            >
                <div>
                    <TextField
                        label="Group Name"
                        value={this.state.groupName}
                        onChange={this.onChangegroupName}
                    />
                    <TextField
                        multiline
                        label="Group Description"
                        value={this.state.groupDescription}
                        onChange={this.onChangegroupDescription}
                    />
                </div>
                <div style={{ marginTop: 10 }}>
                    <PrimaryButton text={this.props.selectedpanelGroup ? "Update Group" : "Add Group"} onClick={() => {
                        this.props.selectedpanelGroup ? this.props.onUpdate(this.props.selectedpanelGroup,
                            { "name": `${this.state.groupName}`, "description": `${this.state.groupDescription}` }) :
                            this.props.onAdd({ "name": `${this.state.groupName}`, "description": `${this.state.groupDescription}` });
                    }} />
                    <PrimaryButton text="Cancel" style={{ marginLeft: 10 }} onClick={this.props.dismissPanelGroup} />
                </div>
            </Panel>
        );
    }

}