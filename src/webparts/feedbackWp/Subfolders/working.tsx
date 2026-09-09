import * as React from 'react';
import styles from './FeedbackWp.module.scss';
import * as strings from 'FeedbackWpWebPartStrings';
import { IFeedbackWpProps } from './IFeedbackWpProps';
import { FeedbackService, IFeedbackPrompt } from '../services/FeedbackService';

type FlowState = 'loading' | 'voting' | 'commentPending' | 'submitted' | 'inactive';
type VoteValue = 'yes' | 'no' | undefined;

interface IFeedbackWpState {
  flow: FlowState;
  prompt: IFeedbackPrompt | undefined;
  vote: VoteValue;
  comment: string;
  isSubmitting: boolean;
  error: string | undefined;
}
/**
 * Feedback web part — collects a Yes/No vote plus an optional comment against
 * the currently active FeedbackPrompts item, and writes each submission to
 * FeedbackResponses. See FeedbackService.ts for the SharePoint REST calls.
 */
export default class FeedbackWp extends React.Component<IFeedbackWpProps, IFeedbackWpState> {
  private service: FeedbackService;

  constructor(props: IFeedbackWpProps) {
    super(props);
    this.service = new FeedbackService(props.context);
    this.state = {
      flow: 'loading',
      prompt: undefined,
      vote: undefined,
      comment: '',
      isSubmitting: false,
      error: undefined
    };
  }

  public componentDidMount(): void {
    if (!this.props.isActive) {
      this.setState({ flow: 'inactive' });
      return;
    }
    this.loadPrompt().catch(() => {
      // loadPrompt() already handles its own errors internally (sets flow: 'inactive'
      // with an error message) — this catch exists only to satisfy no-floating-promises.
    });
  }

  private getErrorMessage = (err: unknown): string => {
    if (err instanceof Error) {
      return err.message;
    }

    return 'Something went wrong.';
  };

  private loadPrompt = async (): Promise<void> => {
    try {
     const prompt = await this.service.getActivePrompt();
      if (!prompt) {
        this.setState({ flow: 'inactive' });
        return;
      }
      this.setState({ flow: 'voting', prompt });
    } catch (err) {
      this.setState({ flow: 'inactive', error: this.getErrorMessage(err) });
    }
  };

  private onVoteClick = (vote: VoteValue): void => {
    this.setState({ vote, flow: 'commentPending' });
  };

  private onCommentChange = (e: React.ChangeEvent<HTMLTextAreaElement>): void => {
    this.setState({ comment: e.target.value });
  };

  private onCancel = (): void => {
    this.setState({ vote: undefined, comment: '', flow: 'voting', error: undefined });
  };

  private onSubmit = async (): Promise<void> => {
    const { prompt, vote, comment } = this.state;
    if (!prompt || !vote) {
      return;
    }

    this.setState({ isSubmitting: true, error: undefined });

    try {
      await this.service.addResponse({
        Title: prompt.Title,
        Comments: comment,
        Like: vote === 'yes',
        PromptId: prompt.Id
      });

      this.setState({ flow: 'submitted', isSubmitting: false });
    } catch (err) {
      this.setState({ isSubmitting: false, error: this.getErrorMessage(err) });
    }
  };

  public render(): React.ReactElement | null {
    const { flow, prompt, vote, comment, isSubmitting, error } = this.state;

    if (flow === 'inactive' || flow === 'loading' || !prompt) {
      // eslint-disable-next-line @rushstack/no-new-null -- see justification above
      return null;
    }

    return (
      <div className={styles.feedbackCard}>
        {flow !== 'submitted' && (
          <div className={styles.headerRow}>
            <div className={styles.iconBadge}>
              <ClipboardIcon />
            </div>
            <div className={styles.questionText}>{prompt.Title}</div>

            <div className={styles.voteButtons}>
              <button
                type="button"
                className={vote === 'yes' ? `${styles.voteBtn} ${styles.voteBtnSelected}` : styles.voteBtn}
                onClick={() => this.onVoteClick('yes')}
                disabled={isSubmitting}
              >
                <ThumbsUpIcon selected={vote === 'yes'} /> {strings.YesButtonLabel}
              </button>
              <button
                type="button"
                className={vote === 'no' ? `${styles.voteBtn} ${styles.voteBtnSelected}` : styles.voteBtn}
                onClick={() => this.onVoteClick('no')}
                disabled={isSubmitting}
              >
                <ThumbsDownIcon selected={vote === 'no'} /> {strings.NoButtonLabel}
              </button>
            </div>
          </div>
        )}

        {flow === 'commentPending' && (
          <div className={styles.commentSection}>
            <textarea
              className={styles.commentBox}
              placeholder={strings.CommentPlaceholder}
              value={comment}
              onChange={this.onCommentChange}
              disabled={isSubmitting}
            />
            <div className={styles.actionButtons}>
              <button
                type="button"
                className={styles.submitBtn}
                onClick={this.onSubmit}
                disabled={isSubmitting}
              >
                {isSubmitting ? strings.SubmittingButtonLabel : strings.SubmitButtonLabel}
              </button>
              <button
                type="button"
                className={styles.cancelBtn}
                onClick={this.onCancel}
                disabled={isSubmitting}
              >
                {strings.CancelButtonLabel}
              </button>
            </div>
            {error && <div className={styles.errorText}>{error}</div>}
          </div>
        )}

        {flow === 'submitted' && (
          <div className={styles.thankYouSection}>
            <div className={styles.checkIcon}>
              <CheckIcon />
            </div>
            <div>
              <span className={styles.thankYouTitle}>{strings.ThankYouTitle}</span>{' '}
              <span className={styles.thankYouSubtitle}>
                {strings.ThankYouSubtitle}
              </span>
            </div>
          </div>
        )}
      </div>
    );
  }
}


function ClipboardIcon(): JSX.Element {
  return (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none">
      <rect x="5" y="4" width="14" height="17" rx="2" stroke="currentColor" strokeWidth="1.6" />
      <rect x="9" y="2.5" width="6" height="3" rx="1" fill="currentColor" />
      <line x1="8" y1="9" x2="16" y2="9" stroke="currentColor" strokeWidth="1.4" />
      <line x1="8" y1="12.5" x2="16" y2="12.5" stroke="currentColor" strokeWidth="1.4" />
      <line x1="8" y1="16" x2="13" y2="16" stroke="currentColor" strokeWidth="1.4" />
    </svg>
  );
}

function ThumbsUpIcon({ selected }: { selected: boolean }): JSX.Element {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
      <path
        d="M7 11v9H4a1 1 0 01-1-1v-7a1 1 0 011-1h3zm0 0l4.5-8a2 2 0 013.6 1.4L14 9h5a2 2 0 012 2l-1.6 7.2A2 2 0 0117.4 20H10a3 3 0 01-3-3v-6z"
        stroke={selected ? '#FFFFFF' : 'currentColor'}
        strokeWidth="1.6"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function ThumbsDownIcon({ selected }: { selected: boolean }): JSX.Element {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
      <path
        d="M17 13V4h3a1 1 0 011 1v7a1 1 0 01-1 1h-3zm0 0l-4.5 8a2 2 0 01-3.6-1.4L10 15H5a2 2 0 01-2-2l1.6-7.2A2 2 0 016.6 4H14a3 3 0 013 3v6z"
        stroke={selected ? '#FFFFFF' : 'currentColor'}
        strokeWidth="1.6"
        strokeLinejoin="round"
      />
    </svg>
  );
}

function CheckIcon(): JSX.Element {
  return (
    <svg width="28" height="28" viewBox="0 0 24 24" fill="none">
      <circle cx="12" cy="12" r="10" stroke="#2E7D32" strokeWidth="1.6" />
      <path d="M8 12.5l2.5 2.5L16 9.5" stroke="#2E7D32" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}
