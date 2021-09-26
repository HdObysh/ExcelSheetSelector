import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";

export interface ProgressProps {
  message: string;
}

export default class Progress extends React.Component<ProgressProps> {
  render() {
    const { message } = this.props;

    return (
      <section className="hdobysh-exaddin__progress ms-u-fadeIn500">
        <Spinner size={SpinnerSize.large} label={message} />
      </section>
    );
  }
}
