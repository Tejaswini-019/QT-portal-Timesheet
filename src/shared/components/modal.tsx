import * as React from 'react';

export interface IModal {
    title: string;
    isOpen: boolean;
    body: string;
    backdrop?: boolean;
    hideSaveBtn?: boolean;
}

interface IBSModalProps {
    title: string;
    html: string;
    backdrop?: boolean;
    display: boolean;
    hideSaveBtn?: boolean;
    onSaveBtnClick?: () => void;
    onCancelBtnClick?: () => void;
    onCloseModal?: () => void;
    footerSaveBtnName?: string;
}

export default class BSModal extends React.Component<IBSModalProps, {}> {

    constructor(props) {
        super(props);
    }

    componentDidUpdate(){
        this.handleModelShow();
    }

    componentDidMount(){
        let __this = this;
        $("#customModal").on("hidden.bs.modal", function () {
            __this.props.onCloseModal();
        });
        this.handleModelShow();
    }

    private handleModelShow() {
        if(this.props.display)
            ($("#customModal") as any).modal('show');
        else
            ($("#customModal") as any).modal('hide');
    }

    render(){
        return(
            <div>
            {/* Modal */}
                <div
                    className="modal fade"
                    id="customModal"
                    tabIndex={-1}
                    role="dialog"
                    data-backdrop= { (this.props.backdrop == undefined || !this.props.backdrop) ? true : "static"}
                    aria-labelledby="exampleModalLabel"
                    aria-hidden="true"
                >
                    <div className="modal-dialog" role="document">
                    <div className="modal-content">
                        <div className="modal-header">
                        <h5 className="modal-title" id="exampleModalLabel">
                            {this.props.title}
                        </h5>
                        <button
                            type="button"
                            className="close"
                            data-dismiss="modal"
                            aria-label="Close"
                        >
                            <span aria-hidden="true">×</span>
                        </button>
                        </div>
                        <div className="modal-body"> 
                            <div dangerouslySetInnerHTML={{__html: this.props.html}}>
                            </div> 
                        </div>
                        <div className="modal-footer">
                        <button
                            type="button"
                            id="bs-cancel"
                            className="btn btn-secondary"
                            data-dismiss="modal"
                            onClick= {this.props.onCancelBtnClick}
                        >
                            Close
                        </button>
                        {
                            ( this.props.hideSaveBtn ) ? null : 
                                    <button type="button" id="bs-save" className="btn btn-primary" onClick={this.props.onSaveBtnClick}>
                                        { (this.props.footerSaveBtnName == undefined) ? "Save Changes" : this.props.footerSaveBtnName }
                                    </button>
                        }
                        </div>
                    </div>
                    </div>
                </div>
            </div>

        )
    }
}