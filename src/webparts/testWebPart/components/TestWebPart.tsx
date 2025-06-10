import * as React from 'react';
import styles from './TestWebPart.module.scss';
import type { ITestWebPartProps } from './ITestWebPartProps';
import { M365Toolkit } from '../../../toolkit/M365Toolkit';
import { PrimaryButton } from '@fluentui/react';

interface ITestWebPartState {
  items: any[];
  nextLink?: string;
  previousLink?: string;
  count?: number;
}

export default class TestWebPart extends React.Component<ITestWebPartProps, ITestWebPartState> {
  constructor(props: ITestWebPartProps) {
    super(props);
    this.state = {
      items: [],
      nextLink: undefined,
      previousLink: undefined,
      count: undefined
    };
  }

  public async componentDidMount(): Promise<void> {
    const temp = await M365Toolkit
      .SPSite(this.props.context)
      .SPWeb()
      .SPList
      .getByTitle('Customers')
      .SPItems
      .select('Title')
      .top(1)
      .get();
    this.setState({ items: temp.value, nextLink: temp['odata.nextLink'] });
  }

  private async fetchMoreItems(url: string): Promise<void> {
    if (this.state.nextLink) {
      const response = await M365Toolkit
        .SPSite(this.props.context)
        .SPWeb()
        .SPList
        .getByUrl(url)
        .SPItems
        .get();
      this.setState((prevState) => ({
        items: [...prevState.items, ...response.value],
        nextLink: response['odata.nextLink'],
        previousLink: response['odata.previousLink'],
        count: response['odata.count']
      }));
    }
  }

  public render(): React.ReactElement<ITestWebPartProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
      <section className={`${styles.testWebPart} ${hasTeamsContext ? styles.teams : ''}`}>
        <h3>Fetched Items (JSON):</h3>
        <pre>
          {JSON.stringify(this.state.items, null, 2)}
        </pre>
        {this.state.nextLink && (
          <PrimaryButton onClick={() => this.fetchMoreItems(this.state.nextLink!)} text="Next Items" />
        )}
        {this.state.previousLink && (
          <PrimaryButton onClick={() => this.fetchMoreItems(this.state.previousLink!)} text="Prevoius Items" />
        )}
        {this.state.count !== undefined && (
          <div>
            <p>Total Count: {this.state.count}</p>
          </div>
        )}
      </section>
    );
  }
}