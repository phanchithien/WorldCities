import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';

import { environment } from '../environments/environment';

export abstract class BaseService<T> {
  constructor(
    protected http: HttpClient
  ) { }

  abstract getData(
    filterQuery: string | null): Observable<ApiResult<T>>;
  abstract put(item: T): Observable<T>;
  }
}
export interface ApiResult<T> {
  pageIndex: number;
  totalCount: number;
  filterColumn: string;
}
