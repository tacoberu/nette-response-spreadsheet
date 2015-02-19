nette-response-spreadsheet
==========================

Spreadsheet Nette responses.

Use:

```php
use Taco\Nette\Application\Responses;

class SomePresenter extends BasePresenter
{
    function actionDefault()
    {
        $data = [
            [ 'George', 'age' => 15, 'grade' => 2, ],
            [ 'Jack', 'age' => 17, 'grade' => 4, ],
            [ 'name' => 'Mary', 'age' => 17, 'grade' => 1, ],
        ];

        $response = new Responses\SpreadsheetResponse($data);
        $this->sendResponse( $response );
    }
}
```

With headers

```php
use Taco\Nette\Application\Responses;

class SomePresenter extends BasePresenter
{
    function actionDefault()
    {
        $headers = [ 'Name', 'Age', 'Grade'];
        $data = [
            [ 'George', 15, 2, ],
            [ 'Jack', 17, 4, ],
            [ 'Mary', 17, 1, ],
        ];

        $response = new Responses\SpreadsheetResponse($data, $headers);
        $this->sendResponse( $response );
    }
}
```

Individual settings example:

```php
use Taco\Nette\Application\Responses;

class SomePresenter extends BasePresenter
{
    function actionDefault()
    {
        $headers = [ 'Name', 'Age', 'Grade'];
        $data = [
            [ 'George', 15, 2, ],
            [ 'Jack', 17, 4, ],
            [ 'Mary', 17, 1, ],
        ];

        $response = new Responses\SpreadsheetResponse($data, $headers);
        $response
            ->setFilename('export')
            ->setTitle('Export');
        $this->sendResponse( $response );
    }
}


```

Multiple sheets example:

```php
use Taco\Nette\Application\Responses;

class SomePresenter extends BasePresenter
{
    function actionDefault()
    {
        $headers1 = [ 'Name', 'Age', 'Grade'];
        $data1 = [
            [ 'George', 15, 2, ],
            [ 'Jack', 17, 4, ],
            [ 'Mary', 17, 1, ],
        ];

        $headers2 = [ 'City', 'Population'];
        $data2 = [
            [ 'Prague', 1.257, ],
            [ 'Paris', 2.211, ],
            [ 'London', 8.308, ],
        ];

        $response = new Responses\SpreadsheetResponse($data1, $headers1);
        $response
            ->setFilename('export')
            ->setTitle('Export');
            ->addSheet($data2, $headers2);
        $this->sendResponse( $response );
    }
}


```
